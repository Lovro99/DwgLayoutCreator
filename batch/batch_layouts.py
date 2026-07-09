#!/usr/bin/env python3
"""
batch_layouts.py — headless batch nad DWG projektima preko accoreconsole.

Uz postojece kreiranje layouta (Korak 2), u ISTOM prolazu po datoteci obavlja
jos 4 opcionalna koraka (plan V2), svaki portan iz interaktivnog LISP-a:
  - polja       upis custom drawing properties iz Excela (SetFieldsValue.lsp)
  - naslovi     upis naslova/mjerila u sastAu blok      (SetLayoutTitles.lsp)
  - sortiranje  numericko sortiranje layout tabova       (TabSort + ele:layout<)
  - export      ispis sortiranih imena layouta u Excel   (ExportLayoutsToExcel.lsp)

Excel se cita/pise BEZ Excela i BEZ vanjskih paketa (xlsx_lite.py, stdlib).
Podaci se pluginu prosljedjuju kroz privremeni JSON + env var LAYOUT_BATCH_JSON.

Za svaki DWG (kad ima izmjena):
  1. sigurnosne provjere (postoji, nije zakljucan, upisiv),
  2. kopija u temp (original se NE dira do uspjesne verifikacije),
  3. create pass:  NETLOAD -> [CREATELAYOUTBATCH][SETFIELDSBATCH][SETTITLESBATCH]
                   [SORTTABSBATCH][UPDATEFIELD] -> QSAVE  (samo ukljuceni koraci),
  4. verify pass:  NETLOAD -> VERIFYBATCH (read-only dump; bez QSAVE),
  5. usporedbe (layouti/polja/naslovi/sortiranje) na KOPIJI,
  6. uspjeh -> original -> <ime>.dwg.bak, kopija -> na mjesto originala,
     neuspjeh -> original NETAKNUT,
  7. export (ako ukljucen): Python sam upise imena u Excel iz verificirane liste,
  8. sazetak: logs/run_<ts>.log + logs/rezultati_<ts>.csv (UTF-8 BOM; stupac 'korak').

SAMO stdlib. Vidi batch/config.json i README.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = SCRIPT_DIR / "templates"
LOG_DIR = SCRIPT_DIR / "logs"

# lokalni moduli (u istom folderu)
sys.path.insert(0, str(SCRIPT_DIR))
import layout_names  # noqa: E402
import xlsx_lite  # noqa: E402

OK, WARN, ERR = "OK", "WARN", "ERR"

# --- markeri iz create passa ---
_RESULT_RE = re.compile(r"RESULT\|\s*created=(\d+)\s+candidates=(\d+)")
_CREATED_RE = re.compile(r"^\s*CREATED\|\s*(.+?)\s*$", re.MULTILINE)
_RESULT_FIELDS_RE = re.compile(r"RESULT\|fields\|added=(\d+)\s+overwritten=(\d+)")
_RESULT_TITLES_RE = re.compile(r"RESULT\|titles\|ok=(\d+)\s+noblock=(\d+)\s+nolayout=(\d+)")
_RESULT_SORT_RE = re.compile(r"RESULT\|sorttabs\|sorted=(\d+)\s+other=(\d+)")
_TITLES_OK_RE = re.compile(r"^\s*TITLES\|ok\|(.*?)\s*$", re.MULTILINE)
_SASTERR_RE = re.compile(r"ERR\|\s*sastAu\.dwg not found")
# Live-progress signali iz plugina (LayoutBuilder ispisuje "[5/5] Done — ... 'ime'").
_FOUND_RE = re.compile(r"Found (\d+) block")
_DONE_RE = re.compile(r"\[5/5\].*?'([^']+)'")
_EDU_RE = re.compile(r"educational", re.IGNORECASE)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def decode_accore(raw: bytes) -> str:
    """accoreconsole zna na pipe ispisati UTF-16LE (znakovi + \\x00). Detektiraj
    i dekodiraj ispravno; inace UTF-8."""
    if raw.count(b"\x00") > len(raw) // 4:
        return raw.decode("utf-16-le", errors="replace")
    return raw.decode("utf-8", errors="replace")


def lisp_path(p: Path | str) -> str:
    """LISP (open ...) i (command NETLOAD ...) traze forward slasheve u stringu."""
    return str(p).replace("\\", "/")


def render_template(tpl: Path, **repl: str) -> str:
    text = tpl.read_text(encoding="ascii")
    for key, val in repl.items():
        text = text.replace(key, val)
    return text


def is_locked(dwg: Path) -> bool:
    """DWG otvoren u AutoCAD-u ostavlja .dwl/.dwl2 lock uz sebe."""
    return dwg.with_suffix(".dwl").exists() or dwg.with_suffix(".dwl2").exists()


def is_writable(dwg: Path) -> bool:
    return os.access(dwg, os.W_OK)


def build_commands(steps: "Steps") -> list[str]:
    """Slozi listu naredbi za create-pass .scr, TOCNO ukljucenim koracima i
    ovim redom (plan §3). UPDATEFIELD ide NAKON svih izmjena (samo ako polja).
    Zadnji je ._QSAVE. Export se NE radi u AutoCAD-u (radi ga Python)."""
    cmds: list[str] = []
    if steps.layouti:
        cmds.append("CREATELAYOUTBATCH")
    if steps.polja:
        cmds.append("SETFIELDSBATCH")
    if steps.naslovi:
        cmds.append("SETTITLESBATCH")
    if steps.sortiranje:
        cmds.append("SORTTABSBATCH")
    if steps.polja:
        # _.UPDATEFIELD _All  + prazna linija (kraj selekcije) — vjeran ekvivalent
        # (command "_.UPDATEFIELD" "_All" "") iz SetFieldsValue.lsp.
        cmds.append("_.UPDATEFIELD")
        cmds.append("_All")
        cmds.append("")
    cmds.append("._QSAVE")
    return cmds


# ---------------------------------------------------------------------------
# config
# ---------------------------------------------------------------------------

@dataclass
class Steps:
    layouti: bool = True
    polja: bool = False
    naslovi: bool = False
    sortiranje: bool = False
    export: bool = False


@dataclass
class Project:
    dwg: Path
    excel: Path | None = None


@dataclass
class Config:
    accoreconsole: Path
    plugin_dll: Path
    sastavnica: Path
    timeout_s: int = 300
    backup: bool = True
    jobs: int = 1
    lang: str = "en-US"
    excel: Path | None = None
    sheet_polja: str = "Podaci"
    sheet_nacrti: str = "Nacrti"
    koraci: Steps = field(default_factory=Steps)
    projekti: list[Project] = field(default_factory=list)


def parse_steps(k: dict | None) -> Steps:
    """koraci izostavljen (ili prazan) -> {layouti: true}, sve ostalo false.
    Ako je prisutan, layouti default True (staro ponasanje), ostali False."""
    if not k:
        return Steps()
    return Steps(
        layouti=bool(k.get("layouti", True)),
        polja=bool(k.get("polja", False)),
        naslovi=bool(k.get("naslovi", False)),
        sortiranje=bool(k.get("sortiranje", False)),
        export=bool(k.get("export", False)),
    )


def resolve_projects(raw: dict, cfg_path: Path) -> list[Project]:
    """Dva nacina: eksplicitna lista 'projekti' (string ILI {dwg, excel}), ili
    'folder' + 'pattern' (glob). Backward compat: stari format (lista stringova)
    radi identicno."""
    projekti: list[Project] = []
    if raw.get("projekti"):
        for item in raw["projekti"]:
            if isinstance(item, str):
                projekti.append(Project(Path(item)))
            elif isinstance(item, dict):
                dwg = item.get("dwg")
                if not dwg:
                    continue
                exc = item.get("excel")
                projekti.append(Project(Path(dwg), Path(exc) if exc else None))
    elif raw.get("folder"):
        folder = Path(raw["folder"])
        pattern = raw.get("pattern", "**/*.dwg")
        projekti = [Project(p) for p in sorted(folder.glob(pattern))]
    else:
        sys.exit(f"config: navedi 'projekti' (lista) ILI 'folder'(+'pattern') u {cfg_path}")
    return [pr for pr in projekti if pr.dwg.suffix.lower() == ".dwg"]


def load_config(path: Path, jobs_override: int | None) -> Config:
    with path.open(encoding="utf-8") as fh:
        raw = json.load(fh)

    def req(key: str) -> str:
        if key not in raw or not str(raw[key]).strip():
            sys.exit(f"config: nedostaje obavezno polje '{key}' u {path}")
        return str(raw[key])

    projekti = resolve_projects(raw, path)
    excel = raw.get("excel")

    cfg = Config(
        accoreconsole=Path(req("accoreconsole")),
        plugin_dll=Path(req("plugin_dll")),
        sastavnica=Path(req("sastavnica")),
        timeout_s=int(raw.get("timeout_s", 300)),
        backup=bool(raw.get("backup", True)),
        jobs=int(jobs_override if jobs_override else raw.get("jobs", 1)),
        lang=str(raw.get("lang", "en-US")),
        excel=Path(excel) if excel else None,
        sheet_polja=str(raw.get("sheet_polja", "Podaci")),
        sheet_nacrti=str(raw.get("sheet_nacrti", "Nacrti")),
        koraci=parse_steps(raw.get("koraci")),
        projekti=projekti,
    )
    return cfg


# ---------------------------------------------------------------------------
# Excel citanje (za korake polja / naslovi)
# ---------------------------------------------------------------------------

def read_polja(excel: Path, sheet: str) -> dict[str, str]:
    """Sheet A=kljuc, B=vrijednost. Redci s praznim A se preskacu (kao LISP
    'prazni redci'); cita se cijeli sheet (bez limita od 200)."""
    grid = xlsx_lite.read_sheet(str(excel), sheet)
    polja: dict[str, str] = {}
    for row in grid:
        key = row[0] if len(row) > 0 else ""
        if key == "":
            continue
        value = row[1] if len(row) > 1 else ""
        polja[key] = value
    return polja


def read_naslovi(excel: Path, sheet: str) -> tuple[list[dict], list[str]]:
    """Sheet A=ime layouta, B=naslov, C=mjerilo. Prazan A -> preskoci; A prisutan
    a B i C oba prazna -> preskoci + WARN; jedan od B/C prazan -> prazan string.
    Vraca (lista naslova, lista upozorenja)."""
    grid = xlsx_lite.read_sheet(str(excel), sheet)
    naslovi: list[dict] = []
    warns: list[str] = []
    for i, row in enumerate(grid, start=1):
        layout = row[0] if len(row) > 0 else ""
        if layout == "":
            continue
        title = row[1] if len(row) > 1 else ""
        mjerilo = row[2] if len(row) > 2 else ""
        if title == "" and mjerilo == "":
            warns.append(f"redak {i}: layout '{layout}' nema podataka (naslov i mjerilo prazni) — preskocen")
            continue
        naslovi.append({"layout": layout, "naslov": title, "mjerilo": mjerilo})
    return naslovi, warns


# ---------------------------------------------------------------------------
# accoreconsole
# ---------------------------------------------------------------------------

@dataclass
class PassResult:
    returncode: int
    stdout: str
    timed_out: bool
    aborted: bool = False


def run_accore(cfg: Config, dwg: Path, scr: Path, env: dict,
               on_line=None, should_abort=None) -> PassResult:
    """Pokreni accoreconsole. Ako je on_line ILI should_abort zadan, streamaj
    stdout (citac thread) da mozemo uzivo emitirati linije (on_line) i/ili
    prekinuti proces (should_abort() -> kill). Inace jednostavno uhvati cijeli
    izlaz na kraju (subprocess.run)."""
    cmd = [str(cfg.accoreconsole), "/i", str(dwg), "/s", str(scr), "/l", cfg.lang]

    if on_line is None and should_abort is None:
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=cfg.timeout_s, env=env)
            return PassResult(proc.returncode, decode_accore(proc.stdout), False)
        except subprocess.TimeoutExpired as exc:
            return PassResult(-1, decode_accore(exc.stdout or b""), True)

    # streaming put: citac u thread-u puni buffer, glavni thread emitira/prekida
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
    raw = bytearray()
    lock = threading.Lock()

    def reader() -> None:
        while True:
            chunk = proc.stdout.read(4096)  # type: ignore[union-attr]
            if not chunk:
                break
            with lock:
                raw.extend(chunk)

    th = threading.Thread(target=reader, daemon=True)
    th.start()

    def snapshot_lines() -> list[str]:
        with lock:
            return decode_accore(bytes(raw)).split("\n")

    start = time.monotonic()
    emitted = 0
    timed_out = False
    aborted = False
    while proc.poll() is None:
        if should_abort is not None and should_abort():
            proc.kill()
            aborted = True
            break
        if time.monotonic() - start > cfg.timeout_s:
            proc.kill()
            timed_out = True
            break
        time.sleep(0.4)
        if on_line is not None:
            complete = snapshot_lines()[:-1]  # zadnja je mozda nedovrsena
            for line in complete[emitted:]:
                on_line(line)
            emitted = len(complete)

    th.join(timeout=2)
    if on_line is not None:
        for line in snapshot_lines()[emitted:]:  # zavrsni flush
            on_line(line)
    with lock:
        text = decode_accore(bytes(raw))
    rc = proc.returncode if proc.returncode is not None else -1
    return PassResult(rc, text, timed_out, aborted)


def parse_create(stdout: str) -> tuple[list[str], int | None, int | None]:
    """Vrati (created_names, created_count, candidates) iz create-pass stdouta."""
    created = [m.group(1) for m in _CREATED_RE.finditer(stdout)]
    rm = _RESULT_RE.search(stdout)
    if rm:
        return created, int(rm.group(1)), int(rm.group(2))
    return created, None, None


def parse_verify(stdout: str) -> tuple[list[tuple[int, str]], dict[str, str], dict[str, tuple[str, str]]]:
    """Parsiraj VERIFYBATCH markere. Vrijednosti mogu sadrzavati '|' -> split s
    ogranicenim brojem splitova (s desna kad treba). Vraca
    (vlayouts=[(taborder,ime)], vprops={kljuc:vrijednost}, vtitles={layout:(title,mjerilo)})."""
    vlayouts: list[tuple[int, str]] = []
    vprops: dict[str, str] = {}
    vtitles: dict[str, tuple[str, str]] = {}
    for raw in stdout.split("\n"):
        line = raw.rstrip("\r\n")
        probe = line.lstrip()
        if probe.startswith("VLAYOUT|"):
            rest = probe[len("VLAYOUT|"):]
            if "|" in rest:
                tab, nm = rest.split("|", 1)
                try:
                    vlayouts.append((int(tab), nm))
                except ValueError:
                    pass
        elif probe.startswith("VPROP|"):
            rest = probe[len("VPROP|"):]
            if "=" in rest:
                k, v = rest.split("=", 1)
                vprops[k] = v
        elif probe.startswith("VTITLE|"):
            rest = probe[len("VTITLE|"):]
            if "|" in rest:
                layout, tail = rest.split("|", 1)   # layout ne sadrzi '|'
                if "|" in tail:
                    title, mjerilo = tail.rsplit("|", 1)  # mjerilo je zadnji
                else:
                    title, mjerilo = tail, ""
                vtitles[layout] = (title, mjerilo)
    return vlayouts, vprops, vtitles


# ---------------------------------------------------------------------------
# verifikacija po koraku (plan §6)
# ---------------------------------------------------------------------------

def verify_sort_order(vlayouts: list[tuple[int, str]]) -> list[str]:
    """Provjeri da su valjani layouti numericki poredani i svi PRIJE nevaljanih."""
    order = [name for _, name in sorted(vlayouts)]
    valid_in_order = [n for n in order if layout_names.valid_layout(n)]
    expected_valid = layout_names.sort_valid_layouts(order)
    errs: list[str] = []
    if valid_in_order != expected_valid:
        errs.append(f"sortiranje: valjani layouti nisu numericki poredani "
                    f"({valid_in_order} != {expected_valid})")
    pos_valid = [i for i, n in enumerate(order) if layout_names.valid_layout(n)]
    pos_invalid = [i for i, n in enumerate(order) if not layout_names.valid_layout(n)]
    if pos_valid and pos_invalid and max(pos_valid) > min(pos_invalid):
        errs.append("sortiranje: nevaljani layout je ispred valjanog (moraju biti iza)")
    return errs


def verify_checks(steps: Steps, created: list[str], polja: dict[str, str],
                  naslovi: list[dict], titles_ok: list[str],
                  vlayouts: list[tuple[int, str]], vprops: dict[str, str],
                  vtitles: dict[str, tuple[str, str]]) -> list[str]:
    """Vrati listu poruka o gresci (prazna = sve OK). Bilo koji mismatch -> ERR,
    original NETAKNUT."""
    errs: list[str] = []
    vnames_ci = {n.lower() for _, n in vlayouts}

    if steps.layouti:
        for c in created:
            if c.lower() not in vnames_ci:
                errs.append(f"layouti: kreirani '{c}' nije u VLAYOUT (nije spremljen)")

    if steps.polja:
        for k, v in polja.items():
            if k not in vprops:
                errs.append(f"polja: kljuc '{k}' nije upisan")
            elif vprops[k] != v:
                errs.append(f"polja: '{k}' = '{vprops[k]}' != poslano '{v}'")

    if steps.naslovi:
        sent = {row["layout"].lower(): (row["naslov"], row["mjerilo"]) for row in naslovi}
        vtitles_ci = {k.lower(): val for k, val in vtitles.items()}
        for layout in titles_ok:
            key = layout.lower()
            if key not in vtitles_ci:
                errs.append(f"naslovi: layout '{layout}' oznacen ok ali nema VTITLE")
            elif key in sent and vtitles_ci[key] != sent[key]:
                errs.append(f"naslovi: '{layout}' VTITLE {vtitles_ci[key]} != poslano {sent[key]}")

    if steps.sortiranje:
        errs.extend(verify_sort_order(vlayouts))

    return errs


# ---------------------------------------------------------------------------
# obrada jedne datoteke
# ---------------------------------------------------------------------------

@dataclass
class Row:
    datoteka: str
    korak: str
    layout: str
    status: str
    poruka: str


def process_file(cfg: Config, project: Project, log: logging.Logger,
                 on_line=None, on_progress=None, should_abort=None) -> list[Row]:
    """Obradi jednu datoteku (svi ukljuceni koraci u istom prolazu). Semantika
    kopija->create->verify->zamjena uz .bak je ocuvana; original se NE dira dok
    verify ne prodje. Export (Excel-write) je jedini korak izvan AutoCAD-a."""
    dwg = project.dwg
    name = dwg.name

    if should_abort is not None and should_abort():
        return [Row(name, "", "", ERR, "preskoceno (prekid prije obrade)")]
    if not dwg.is_file():
        return [Row(name, "", "", ERR, "datoteka ne postoji")]
    if is_locked(dwg):
        return [Row(name, "", "", ERR, "zakljucana (.dwl/.dwl2 postoji) — otvorena u AutoCAD-u")]
    if not is_writable(dwg):
        return [Row(name, "", "", ERR, "read-only / nema prava pisanja")]

    steps = cfg.koraci
    # efektivni koraci (excel koraci se mogu iskljuciti ako Excel nedostupan / nescitljiv)
    eff = Steps(layouti=steps.layouti, polja=steps.polja, naslovi=steps.naslovi,
                sortiranje=steps.sortiranje, export=steps.export)
    excel_path = project.excel or cfg.excel

    rows: list[Row] = []

    # --- preflight Excela za korake polja/naslovi/export ---
    if (eff.polja or eff.naslovi or eff.export):
        if excel_path is None or not Path(excel_path).is_file():
            for korak, enabled in (("polja", eff.polja), ("naslovi", eff.naslovi), ("export", eff.export)):
                if enabled:
                    rows.append(Row(name, korak, "", ERR,
                                    f"Excel nedostupan ({excel_path}) — korak preskocen "
                                    f"(layouti/sortiranje se svejedno izvrsavaju)"))
            eff.polja = eff.naslovi = eff.export = False

    # --- citanje Excela (polja / naslovi) ---
    polja: dict[str, str] = {}
    naslovi: list[dict] = []
    if eff.polja:
        try:
            polja = read_polja(Path(excel_path), cfg.sheet_polja)
        except Exception as exc:  # noqa: BLE001
            rows.append(Row(name, "polja", "", ERR, f"citanje sheeta '{cfg.sheet_polja}': {exc}"))
            eff.polja = False
    if eff.naslovi:
        try:
            naslovi, nwarns = read_naslovi(Path(excel_path), cfg.sheet_nacrti)
            for w in nwarns:
                rows.append(Row(name, "naslovi", "", WARN, w))
        except Exception as exc:  # noqa: BLE001
            rows.append(Row(name, "naslovi", "", ERR, f"citanje sheeta '{cfg.sheet_nacrti}': {exc}"))
            eff.naslovi = False

    modifies = eff.layouti or eff.polja or eff.naslovi or eff.sortiranje
    need_verify = modifies or eff.export
    if not need_verify:
        if not any(r.status == ERR for r in rows):
            rows.append(Row(name, "", "", OK, "nema ukljucenih koraka za ovu datoteku"))
        return rows

    only_layouti = (eff.layouti and not (eff.polja or eff.naslovi or eff.sortiranje or eff.export))

    work_dir = Path(tempfile.mkdtemp(prefix="batch_layout_"))
    try:
        env = os.environ.copy()
        env["LAYOUT_SAST_PATH"] = str(cfg.sastavnica)

        # JSON data file (samo za polja/naslovi komande)
        if eff.polja or eff.naslovi:
            data: dict = {}
            if eff.polja:
                data["polja"] = polja
            if eff.naslovi:
                data["naslovi"] = naslovi
            json_path = work_dir / "batch_data.json"
            json_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            env["LAYOUT_BATCH_JSON"] = str(json_path)

        created: list[str] = []
        created_n: int | None = None
        candidates: int | None = None
        create_stdout = ""

        # --- create pass (samo ako ima izmjena) ---
        if modifies:
            copy = work_dir / dwg.name
            shutil.copy2(dwg, copy)

            create_scr = work_dir / "create.scr"
            create_scr.write_text(
                render_template(TEMPLATE_DIR / "create_layouts.scr.tpl",
                                __DLL_PATH__=lisp_path(cfg.plugin_dll),
                                __COMMANDS__="\n".join(build_commands(eff))),
                encoding="ascii",
            )

            line_handler = None
            if on_line is not None or on_progress is not None:
                pstate = {"done": 0, "total": None}

                def line_handler(line: str, _s=pstate) -> None:
                    fm = _FOUND_RE.search(line)
                    if fm:
                        _s["total"] = int(fm.group(1))
                    dm = _DONE_RE.search(line)
                    if dm:
                        _s["done"] += 1
                        if on_progress is not None:
                            on_progress(_s["done"], _s["total"], dm.group(1))
                    if on_line is not None:
                        on_line(line)

            res = run_accore(cfg, copy, create_scr, env,
                             on_line=line_handler, should_abort=should_abort)
            create_stdout = res.stdout
            log.info("[%s] create exit=%s timeout=%s", name, res.returncode, res.timed_out)

            if res.aborted:
                return rows + [Row(name, "", "", ERR, "prekinuto — original netaknut")]
            if res.timed_out:
                log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
                return rows + [Row(name, "", "", ERR, f"timeout > {cfg.timeout_s}s — original netaknut")]
            if _SASTERR_RE.search(res.stdout):
                return rows + [Row(name, "layouti", "", ERR, "sastAu.dwg nedostupan — original netaknut")]
            if res.returncode != 0:
                log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
                return rows + [Row(name, "", "", ERR, f"accoreconsole exit code {res.returncode}")]

            if eff.layouti:
                created, created_n, candidates = parse_create(res.stdout)
                if created_n is None:
                    log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
                    return rows + [Row(name, "layouti", "", ERR,
                                       "nema RESULT| u ispisu (NETLOAD/CREATELAYOUTBATCH nije prosao?)")]
                if created_n > 0 and not created:
                    return rows + [Row(name, "layouti", "", ERR,
                                       "RESULT| created>0 bez CREATED| imena — rebuildaj LayoutCreatorCore (stari DLL?)")]

            # backward-compat kratki put: samo layouti, 0 novih -> original netaknut
            if only_layouti and created_n == 0:
                return rows + [Row(name, "layouti", "", OK,
                                   f"0 novih layouta (idempotentno; kandidata={candidates})")]

            verify_target = copy
        else:
            copy = None
            verify_target = dwg   # samo export: verify na originalu (read-only)

        # --- verify pass (VERIFYBATCH, read-only) ---
        verify_scr = work_dir / "verify.scr"
        verify_scr.write_text(
            render_template(TEMPLATE_DIR / "verify_layouts.scr.tpl",
                            __DLL_PATH__=lisp_path(cfg.plugin_dll)),
            encoding="ascii",
        )
        vres = run_accore(cfg, verify_target, verify_scr, env, should_abort=should_abort)
        log.info("[%s] verify exit=%s", name, vres.returncode)

        if vres.aborted:
            return rows + [Row(name, "verify", "", ERR, "prekinuto tijekom verifikacije — original netaknut")]
        if "RESULT|verify|done" not in vres.stdout:
            log.warning("[%s] verify stdout (zadnjih 2000 zn.):\n%s", name, vres.stdout[-2000:])
            return rows + [Row(name, "verify", "", ERR,
                               "VERIFYBATCH nije dao ocekivani ispis — original netaknut (stari DLL?)")]

        vlayouts, vprops, vtitles = parse_verify(vres.stdout)

        # --- usporedbe (samo kad ima izmjena) ---
        if modifies:
            titles_ok = _TITLES_OK_RE.findall(create_stdout)
            errs = verify_checks(eff, created, polja, naslovi, titles_ok, vlayouts, vprops, vtitles)
            if errs:
                log.warning("[%s] verify FAIL:\n  %s\n--- verify stdout ---\n%s",
                            name, "\n  ".join(errs), vres.stdout[-2000:])
                return rows + [Row(name, "verify", "", ERR,
                                   "verify FAIL: " + "; ".join(errs) + " — original netaknut")]

            # --- uspjeh: zamijeni original (uz .bak) ---
            try:
                if cfg.backup:
                    backup = dwg.with_suffix(dwg.suffix + ".bak")
                    os.replace(dwg, backup)
                else:
                    os.remove(dwg)
                shutil.move(str(copy), str(dwg))
            except OSError as exc:
                return rows + [Row(name, "", "", ERR, f"zamjena originala nije uspjela: {exc} — provjeri .bak")]

            rows.extend(build_step_rows(name, eff, created, created_n, candidates, create_stdout))

        # --- export (Python, NAKON uspjesnog verifyja/zamjene) ---
        if eff.export:
            rows.extend(do_export(name, excel_path, cfg.sheet_nacrti, vlayouts))

        if not rows:
            rows.append(Row(name, "", "", OK, "obradjeno bez izmjena"))
        return rows

    except Exception as exc:  # noqa: BLE001 - zadnja linija obrane; original netaknut
        return rows + [Row(name, "", "", ERR, f"neocekivana greska: {exc} — original netaknut")]
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def build_step_rows(name: str, eff: Steps, created: list[str], created_n: int | None,
                    candidates: int | None, stdout: str) -> list[Row]:
    """Sazetak-redci po koraku iz create-pass markera (nakon uspjeha)."""
    rows: list[Row] = []

    if eff.layouti:
        for c in created:
            rows.append(Row(name, "layouti", c, OK, "kreiran i verificiran"))
        if created_n == 0:
            rows.append(Row(name, "layouti", "", OK, f"0 novih layouta (kandidata={candidates})"))
        elif candidates is not None and created_n is not None and created_n < candidates:
            rows.append(Row(name, "layouti", "", WARN,
                            f"kreirano {created_n}/{candidates} (ostalo preskoceno: X/postoji/custom bez templatea)"))

    if eff.polja:
        fm = _RESULT_FIELDS_RE.search(stdout)
        if fm:
            rows.append(Row(name, "polja", "", OK,
                            f"polja: dodano {fm.group(1)}, prepisano {fm.group(2)}"))
        else:
            rows.append(Row(name, "polja", "", WARN, "nema RESULT|fields| u ispisu"))

    if eff.naslovi:
        tm = _RESULT_TITLES_RE.search(stdout)
        if tm:
            ok_n, nob, nol = tm.group(1), tm.group(2), tm.group(3)
            status = OK if (nob == "0" and nol == "0") else WARN
            rows.append(Row(name, "naslovi", "", status,
                            f"naslovi: ok={ok_n} noblock={nob} nolayout={nol}"))
        else:
            rows.append(Row(name, "naslovi", "", WARN, "nema RESULT|titles| u ispisu"))

    if eff.sortiranje:
        sm = _RESULT_SORT_RE.search(stdout)
        if sm:
            rows.append(Row(name, "sortiranje", "", OK,
                            f"sortiranje: sorted={sm.group(1)} other={sm.group(2)}"))
        else:
            rows.append(Row(name, "sortiranje", "", WARN, "nema RESULT|sorttabs| u ispisu"))

    if _EDU_RE.search(stdout):
        rows.append(Row(name, "", "", WARN,
                        "educational stamp detektiran u crtezu — PDF-ovi ce nositi vodeni zig, provjeri rucno"))
    return rows


def do_export(name: str, excel_path: Path | None, sheet: str,
              vlayouts: list[tuple[int, str]]) -> list[Row]:
    """Export korak (§4.4): sortirana valjana imena layouta -> stupac A sheeta.
    Radi ga Python (bez AutoCAD-a) iz verificirane liste. DWG izmjene ostaju
    valjane i ako export padne (jedini Excel-write korak)."""
    rows: list[Row] = []
    order = [nm for _, nm in sorted(vlayouts)]
    valid_sorted = layout_names.sort_valid_layouts(order)
    try:
        warns = xlsx_lite.write_column_a(str(excel_path), sheet, valid_sorted)
        rows.append(Row(name, "export", "", OK,
                        f"export: upisano {len(valid_sorted)} imena u stupac A sheeta '{sheet}'"))
        for w in warns:
            rows.append(Row(name, "export", "", WARN, f"export: {w}"))
    except Exception as exc:  # noqa: BLE001
        rows.append(Row(name, "export", "", ERR,
                        f"export u Excel nije uspio: {exc} (DWG izmjene ostaju valjane)"))
    return rows


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def preflight(cfg: Config) -> None:
    problems = []
    if not cfg.accoreconsole.is_file():
        problems.append(f"accoreconsole ne postoji: {cfg.accoreconsole}")
    if not cfg.plugin_dll.is_file():
        problems.append(f"plugin_dll ne postoji: {cfg.plugin_dll} (jesi li buildao LayoutCreatorCore?)")
    if cfg.koraci.layouti and not cfg.sastavnica.is_file():
        problems.append(f"sastavnica ne postoji: {cfg.sastavnica} (potrebna za korak 'layouti')")
    if not cfg.projekti:
        problems.append("nema projekata za obradu (provjeri 'projekti' ili 'folder'/'pattern')")
    if problems:
        sys.exit("PREFLIGHT greske:\n  - " + "\n  - ".join(problems))


def setup_logging(ts: str, to_stdout: bool = True) -> logging.Logger:
    """Logger s file handlerom (run_<ts>.log). CLI dodaje i stdout handler;
    GUI ga izostavi (pythonw nema konzolu) i sam prikaci svoj queue handler."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log = logging.getLogger("batch_layouts")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fh = logging.FileHandler(LOG_DIR / f"run_{ts}.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)
    if to_stdout and sys.stdout is not None:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(logging.Formatter("%(message)s"))
        log.addHandler(ch)
    return log


def run_log_path(ts: str) -> Path:
    return LOG_DIR / f"run_{ts}.log"


def read_config_raw(path: Path) -> dict:
    """Ucitaj config.json kao dict (za GUI). Isti format koji cita load_config."""
    with path.open(encoding="utf-8") as fh:
        return json.load(fh)


def write_config(path: Path, data: dict) -> None:
    """Spremi config.json (za GUI) u isti format koji cita CLI load_config."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


def write_csv(ts: str, rows: list[Row]) -> Path:
    csv_path = LOG_DIR / f"rezultati_{ts}.csv"
    # UTF-8 s BOM (utf-8-sig) da Excel ispravno prikaze c/c/s/d/z
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["datoteka", "korak", "layout", "status", "poruka"])
        for r in rows:
            writer.writerow([r.datoteka, r.korak, r.layout, r.status, r.poruka])
    return csv_path


def _steps_summary(steps: Steps) -> str:
    on = [k for k in ("layouti", "polja", "naslovi", "sortiranje", "export") if getattr(steps, k)]
    return ", ".join(on) if on else "(nijedan)"


def main() -> None:
    ap = argparse.ArgumentParser(description="Batch kreiranje/obrada AutoCAD layouta (headless).")
    ap.add_argument("--config", type=Path, default=SCRIPT_DIR / "config.json")
    ap.add_argument("--dry-run", action="store_true", help="samo ispisi sto bi se radilo (+ generirani .scr)")
    ap.add_argument("--jobs", type=int, default=None, help="paralelnih accoreconsole procesa")
    args = ap.parse_args()

    if not args.config.is_file():
        sys.exit(f"config ne postoji: {args.config}")

    cfg = load_config(args.config, args.jobs)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log = setup_logging(ts)

    log.info("Config: %s", args.config)
    log.info("accoreconsole: %s", cfg.accoreconsole)
    log.info("plugin_dll:    %s", cfg.plugin_dll)
    log.info("sastavnica:    %s", cfg.sastavnica)
    log.info("koraci:        %s", _steps_summary(cfg.koraci))
    if cfg.excel:
        log.info("excel:         %s  (polja='%s', nacrti='%s')",
                 cfg.excel, cfg.sheet_polja, cfg.sheet_nacrti)
    log.info("projekata:     %d | jobs=%d | timeout=%ds | backup=%s",
             len(cfg.projekti), cfg.jobs, cfg.timeout_s, cfg.backup)

    if args.dry_run:
        log.info("\n--- DRY RUN (nista se ne mijenja) ---")
        log.info("Generirani create-pass .scr sadrzaj (koraci: %s):", _steps_summary(cfg.koraci))
        log.info("  (command \"._NETLOAD\" \"%s\")", lisp_path(cfg.plugin_dll))
        for c in build_commands(cfg.koraci):
            log.info("  %s", c if c else "(prazna linija = kraj selekcije)")
        log.info("")
        for pr in cfg.projekti:
            flags = []
            if not pr.dwg.is_file():
                flags.append("NEMA")
            elif is_locked(pr.dwg):
                flags.append("ZAKLJUCAN")
            elif not is_writable(pr.dwg):
                flags.append("READ-ONLY")
            xls = pr.excel or cfg.excel
            xnote = "" if not (cfg.koraci.polja or cfg.koraci.naslovi or cfg.koraci.export) \
                else (f"  excel={xls}" if xls else "  excel=NEMA")
            log.info("  %s  %s%s", pr.dwg, " ".join(flags) if flags else "OK", xnote)
        return

    preflight(cfg)

    all_rows: list[Row] = []
    total = len(cfg.projekti)
    if cfg.jobs > 1:
        # paralelno: live \r progres bi se ispreplitao, pa samo start/kraj po datoteci
        with ThreadPoolExecutor(max_workers=cfg.jobs) as ex:
            futures = {ex.submit(process_file, cfg, pr, log): pr for pr in cfg.projekti}
            done = 0
            for fut in as_completed(futures):
                done += 1
                rows = fut.result()
                all_rows.extend(rows)
                pr = futures[fut]
                log.info("[%d/%d] gotovo: %s", done, total, pr.dwg.name)
    else:
        for i, pr in enumerate(cfg.projekti, 1):
            log.info("[%d/%d] %s — obrada (%s)...", i, total, pr.dwg.name, _steps_summary(cfg.koraci))

            def cli_progress(done: int, tot, layout: str, _n=pr.dwg.name) -> None:
                t = f"/{tot}" if tot else ""
                sys.stdout.write(f"\r    {_n}: {done}{t} layouta  ({layout})        ")
                sys.stdout.flush()

            all_rows.extend(process_file(cfg, pr, log, on_progress=cli_progress))
            sys.stdout.write("\n")
            sys.stdout.flush()

    csv_path = write_csv(ts, all_rows)

    n_ok = sum(1 for r in all_rows if r.status == OK)
    n_warn = sum(1 for r in all_rows if r.status == WARN)
    n_err = sum(1 for r in all_rows if r.status == ERR)
    log.info("\n=== SAZETAK === OK=%d WARN=%d ERR=%d", n_ok, n_warn, n_err)
    log.info("CSV: %s", csv_path)
    log.info("Log: %s", LOG_DIR / f"run_{ts}.log")

    # exit code != 0 ako je bilo gresaka (za CI / batch skripte)
    sys.exit(1 if n_err else 0)


if __name__ == "__main__":
    main()
