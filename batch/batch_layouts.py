#!/usr/bin/env python3
"""
batch_layouts.py — headless batch kreiranje paperspace layouta preko vise DWG
projekata (Korak 2 iz IMPLEMENTACIJA_PLAN.md).

Za svaki DWG:
  1. sigurnosne provjere (postoji, nije zakljucan .dwl/.dwl2, upisiv),
  2. kopija u temp (original se NE dira do uspjesne verifikacije),
  3. accoreconsole create pass:  NETLOAD LayoutCreatorCore -> CREATELAYOUTBATCH -> QSAVE,
     (putanja sastavnice se salje kroz env var LAYOUT_SAST_PATH),
  4. accoreconsole verify pass:  ispis (ACAD_LAYOUT rjecnik) imena layouta u txt,
  5. usporedba: svako CREATED| ime iz create passa mora biti u verify listi,
  6. uspjeh -> original -> <ime>.dwg.bak, kopija -> na mjesto originala,
     neuspjeh -> original NETAKNUT,
  7. sazetak: logs/run_<timestamp>.log + logs/rezultati_<timestamp>.csv (UTF-8 BOM).

SAMO stdlib. Parametri layouta (format, broj) ostaju u DWG blokovima — config
samo orkestrira (koje datoteke, alati, sastavnica, timeout). Vidi batch/config.json.

Pokretanje (Windows, AutoCAD 2025 s accoreconsole):
    python batch\\batch_layouts.py --config batch\\config.json
    python batch\\batch_layouts.py --config batch\\config.json --dry-run
    python batch\\batch_layouts.py --config batch\\config.json --jobs 3
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
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = SCRIPT_DIR / "templates"
LOG_DIR = SCRIPT_DIR / "logs"

OK, WARN, ERR = "OK", "WARN", "ERR"

_RESULT_RE = re.compile(r"RESULT\|\s*created=(\d+)\s+candidates=(\d+)")
_CREATED_RE = re.compile(r"^\s*CREATED\|\s*(.+?)\s*$", re.MULTILINE)
_SASTERR_RE = re.compile(r"ERR\|\s*sastAu\.dwg not found")


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


# ---------------------------------------------------------------------------
# config
# ---------------------------------------------------------------------------

@dataclass
class Config:
    accoreconsole: Path
    plugin_dll: Path
    sastavnica: Path
    timeout_s: int = 300
    backup: bool = True
    jobs: int = 1
    lang: str = "en-US"
    projekti: list[Path] = field(default_factory=list)


def load_config(path: Path, jobs_override: int | None) -> Config:
    with path.open(encoding="utf-8") as fh:
        raw = json.load(fh)

    def req(key: str) -> str:
        if key not in raw or not str(raw[key]).strip():
            sys.exit(f"config: nedostaje obavezno polje '{key}' u {path}")
        return str(raw[key])

    projekti = resolve_projects(raw, path)

    cfg = Config(
        accoreconsole=Path(req("accoreconsole")),
        plugin_dll=Path(req("plugin_dll")),
        sastavnica=Path(req("sastavnica")),
        timeout_s=int(raw.get("timeout_s", 300)),
        backup=bool(raw.get("backup", True)),
        jobs=int(jobs_override if jobs_override else raw.get("jobs", 1)),
        lang=str(raw.get("lang", "en-US")),
        projekti=projekti,
    )
    return cfg


def resolve_projects(raw: dict, cfg_path: Path) -> list[Path]:
    """Dva nacina: eksplicitna lista 'projekti', ili 'folder' + 'pattern' (glob)."""
    projekti: list[Path] = []
    if raw.get("projekti"):
        projekti = [Path(p) for p in raw["projekti"]]
    elif raw.get("folder"):
        folder = Path(raw["folder"])
        pattern = raw.get("pattern", "**/*.dwg")
        projekti = sorted(folder.glob(pattern))
    else:
        sys.exit(f"config: navedi 'projekti' (lista) ILI 'folder'(+'pattern') u {cfg_path}")
    # preskoci vec kreirane .bak i ODA temp datoteke
    return [p for p in projekti if p.suffix.lower() == ".dwg"]


# ---------------------------------------------------------------------------
# accoreconsole
# ---------------------------------------------------------------------------

@dataclass
class PassResult:
    returncode: int
    stdout: str
    timed_out: bool


def run_accore(cfg: Config, dwg: Path, scr: Path, env: dict) -> PassResult:
    cmd = [str(cfg.accoreconsole), "/i", str(dwg), "/s", str(scr), "/l", cfg.lang]
    try:
        proc = subprocess.run(
            cmd, capture_output=True, timeout=cfg.timeout_s, env=env
        )
        return PassResult(proc.returncode, decode_accore(proc.stdout), False)
    except subprocess.TimeoutExpired as exc:
        out = decode_accore(exc.stdout or b"")
        return PassResult(-1, out, True)


def parse_create(stdout: str) -> tuple[list[str], int | None, int | None]:
    """Vrati (created_names, created_count, candidates) iz create-pass stdouta."""
    created = [m.group(1) for m in _CREATED_RE.finditer(stdout)]
    rm = _RESULT_RE.search(stdout)
    if rm:
        return created, int(rm.group(1)), int(rm.group(2))
    return created, None, None


# ---------------------------------------------------------------------------
# obrada jedne datoteke
# ---------------------------------------------------------------------------

@dataclass
class Row:
    datoteka: str
    layout: str
    status: str
    poruka: str


def process_file(cfg: Config, dwg: Path, log: logging.Logger) -> list[Row]:
    name = dwg.name

    if not dwg.is_file():
        return [Row(name, "", ERR, "datoteka ne postoji")]
    if is_locked(dwg):
        return [Row(name, "", ERR, "zakljucana (.dwl/.dwl2 postoji) — otvorena u AutoCAD-u")]
    if not is_writable(dwg):
        return [Row(name, "", ERR, "read-only / nema prava pisanja")]

    work_dir = Path(tempfile.mkdtemp(prefix="batch_layout_"))
    copy = work_dir / dwg.name
    try:
        shutil.copy2(dwg, copy)

        env = os.environ.copy()
        env["LAYOUT_SAST_PATH"] = str(cfg.sastavnica)

        # --- create pass ---
        create_scr = work_dir / "create.scr"
        create_scr.write_text(
            render_template(TEMPLATE_DIR / "create_layouts.scr.tpl",
                            __DLL_PATH__=lisp_path(cfg.plugin_dll)),
            encoding="ascii",
        )
        res = run_accore(cfg, copy, create_scr, env)
        log.info("[%s] create exit=%s timeout=%s", name, res.returncode, res.timed_out)

        if res.timed_out:
            log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
            return [Row(name, "", ERR, f"timeout > {cfg.timeout_s}s — original netaknut")]
        if _SASTERR_RE.search(res.stdout):
            return [Row(name, "", ERR, "sastAu.dwg nedostupan — original netaknut")]
        if res.returncode != 0:
            log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
            return [Row(name, "", ERR, f"accoreconsole exit code {res.returncode}")]

        created, created_n, candidates = parse_create(res.stdout)
        if created_n is None:
            log.warning("[%s] create stdout (zadnjih 2000 zn.):\n%s", name, res.stdout[-2000:])
            return [Row(name, "", ERR,
                        "nema RESULT| u ispisu (NETLOAD/CREATELAYOUTBATCH nije prosao?)")]

        if created_n > 0 and not created:
            # RESULT| kaze da je nesto kreirano, ali nema nijednog CREATED| imena
            # -> gotovo sigurno stari DLL. Ne diraj original (inace lazna verifikacija).
            return [Row(name, "", ERR,
                        "RESULT| created>0 bez CREATED| imena — rebuildaj LayoutCreatorCore (stari DLL?)")]

        if created_n == 0:
            # idempotentno: nista novo (svi vec postoje / preskoceni) -> original netaknut
            return [Row(name, "", OK,
                        f"0 novih layouta (idempotentno; kandidata={candidates})")]

        # --- verify pass (na kopiji, PRIJE zamjene originala) ---
        out_txt = work_dir / "layouts.txt"
        verify_scr = work_dir / "verify.scr"
        verify_scr.write_text(
            render_template(TEMPLATE_DIR / "verify_layouts.scr.tpl",
                            __OUT_PATH__=lisp_path(out_txt)),
            encoding="ascii",
        )
        vres = run_accore(cfg, copy, verify_scr, env)
        log.info("[%s] verify exit=%s", name, vres.returncode)

        if not out_txt.exists():
            return [Row(name, "", ERR, "verify pass nije zapisao layouts.txt — original netaknut")]
        persisted = set(out_txt.read_text(encoding="utf-8", errors="replace").split("\n"))
        persisted = {s.strip() for s in persisted if s.strip()}

        missing = [c for c in created if c not in persisted]
        if missing:
            return [Row(name, ", ".join(missing), ERR,
                        f"verify FAIL: {len(missing)} layouta nije spremljeno — original netaknut")]

        # --- uspjeh: zamijeni original (uz .bak) ---
        try:
            if cfg.backup:
                backup = dwg.with_suffix(dwg.suffix + ".bak")
                os.replace(dwg, backup)          # original -> .bak (isti dir)
            else:
                os.remove(dwg)
            shutil.move(str(copy), str(dwg))     # kopija -> original (cross-fs ok)
        except OSError as exc:
            return [Row(name, "", ERR, f"zamjena originala nije uspjela: {exc} — provjeri .bak")]

        rows = [Row(name, c, OK, "kreiran i verificiran") for c in created]
        if candidates is not None and created_n < candidates:
            rows.append(Row(name, "", WARN,
                            f"kreirano {created_n}/{candidates} (ostalo preskoceno: X/postoji/custom bez templatea)"))
        return rows

    except Exception as exc:  # noqa: BLE001 - zadnja linija obrane; original netaknut
        return [Row(name, "", ERR, f"neocekivana greska: {exc} — original netaknut")]
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def preflight(cfg: Config) -> None:
    problems = []
    if not cfg.accoreconsole.is_file():
        problems.append(f"accoreconsole ne postoji: {cfg.accoreconsole}")
    if not cfg.plugin_dll.is_file():
        problems.append(f"plugin_dll ne postoji: {cfg.plugin_dll} (jesi li buildao LayoutCreatorCore?)")
    if not cfg.sastavnica.is_file():
        problems.append(f"sastavnica ne postoji: {cfg.sastavnica}")
    if not cfg.projekti:
        problems.append("nema projekata za obradu (provjeri 'projekti' ili 'folder'/'pattern')")
    if problems:
        sys.exit("PREFLIGHT greske:\n  - " + "\n  - ".join(problems))


def setup_logging(ts: str) -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log = logging.getLogger("batch_layouts")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fh = logging.FileHandler(LOG_DIR / f"run_{ts}.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(fh)
    log.addHandler(ch)
    return log


def write_csv(ts: str, rows: list[Row]) -> Path:
    csv_path = LOG_DIR / f"rezultati_{ts}.csv"
    # UTF-8 s BOM (utf-8-sig) da Excel ispravno prikaze c/c/s/d/z
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["datoteka", "layout", "status", "poruka"])
        for r in rows:
            writer.writerow([r.datoteka, r.layout, r.status, r.poruka])
    return csv_path


def main() -> None:
    ap = argparse.ArgumentParser(description="Batch kreiranje AutoCAD layouta (headless).")
    ap.add_argument("--config", type=Path, default=SCRIPT_DIR / "config.json")
    ap.add_argument("--dry-run", action="store_true", help="samo ispisi sto bi se radilo")
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
    log.info("projekata:     %d | jobs=%d | timeout=%ds | backup=%s",
             len(cfg.projekti), cfg.jobs, cfg.timeout_s, cfg.backup)

    if args.dry_run:
        log.info("\n--- DRY RUN (nista se ne mijenja) ---")
        for p in cfg.projekti:
            flags = []
            if not p.is_file():
                flags.append("NEMA")
            elif is_locked(p):
                flags.append("ZAKLJUCAN")
            elif not is_writable(p):
                flags.append("READ-ONLY")
            log.info("  %s  %s", p, " ".join(flags) if flags else "OK")
        return

    preflight(cfg)

    all_rows: list[Row] = []
    if cfg.jobs > 1:
        with ThreadPoolExecutor(max_workers=cfg.jobs) as ex:
            futures = {ex.submit(process_file, cfg, p, log): p for p in cfg.projekti}
            for fut in as_completed(futures):
                all_rows.extend(fut.result())
    else:
        for p in cfg.projekti:
            all_rows.extend(process_file(cfg, p, log))

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
