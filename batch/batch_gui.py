#!/usr/bin/env python3
"""
Batch Layouti — GUI
===================
Grafičko sučelje za `batch_layouts.py` (headless batch kreiranje AutoCAD layouta),
da se ne mora ručno uređivati `config.json`.

Samo standardna biblioteka (tkinter + ttk). Sva logika obrade dolazi iz
`batch_layouts.py` — GUI je samo tanki sloj: odabir datoteka, postavke, i live
prikaz napretka. Semantika obrade (kopija → create → verify → zamjena uz .bak)
je NEPROMIJENJENA.

Pokretanje:
    python batch\\batch_gui.py
"""

from __future__ import annotations

import logging
import os
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- import domenske logike (batch_layouts.py je u istom folderu) ---
sys.path.insert(0, str(Path(__file__).resolve().parent))
import batch_layouts as core  # noqa: E402

STATE_FILE = core.SCRIPT_DIR / ".batch_gui_state.json"
DEFAULT_CONFIG = core.SCRIPT_DIR / "config.json"
DEFAULT_PATTERN = "**/*.dwg"


# ---------------------------------------------------------------------------
# pomocno
# ---------------------------------------------------------------------------

def file_status(p: Path) -> str:
    if not p.is_file():
        return "NE POSTOJI"
    if core.is_locked(p):
        return "ZAKLJUCAN"
    if not core.is_writable(p):
        return "READ-ONLY"
    return "OK"


class QueueLogHandler(logging.Handler):
    """logging Handler koji formatirane zapise gura u thread-safe queue."""

    def __init__(self, q: "queue.Queue") -> None:
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        self.q.put(("log", self.format(record)))


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class BatchGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Batch Layouti — DwgLayoutCreator")
        root.geometry("880x720")
        root.minsize(760, 600)

        # projekti kao core.Project (cuva i per-projektni excel iz configa)
        self.files: list[core.Project] = []
        self.q: "queue.Queue" = queue.Queue()
        self.abort = threading.Event()
        self.worker: threading.Thread | None = None
        self.last_csv: Path | None = None
        self.last_log: Path | None = None

        self.var_accore = tk.StringVar()
        self.var_dll = tk.StringVar()
        self.var_sast = tk.StringVar()
        self.var_timeout = tk.IntVar(value=300)
        self.var_jobs = tk.IntVar(value=1)
        self.var_backup = tk.BooleanVar(value=True)
        self.var_pattern = tk.StringVar(value=DEFAULT_PATTERN)
        self.var_lang = "en-US"

        # koraci (plan V2) — layouti default ON, ostali OFF (staro ponasanje)
        self.var_k_layouti = tk.BooleanVar(value=True)
        self.var_k_polja = tk.BooleanVar(value=False)
        self.var_k_naslovi = tk.BooleanVar(value=False)
        self.var_k_sortiranje = tk.BooleanVar(value=False)
        self.var_k_export = tk.BooleanVar(value=False)
        self.var_excel = tk.StringVar()
        self.var_sheet_polja = tk.StringVar(value="Podaci")
        self.var_sheet_nacrti = tk.StringVar(value="Nacrti")

        self._build_ui()
        self._load_initial_config()
        self.root.after(120, self._poll_queue)

    # ---- izgradnja UI-a ----
    def _build_ui(self) -> None:
        pad = {"padx": 6, "pady": 3}

        # Postavke
        fs = ttk.LabelFrame(self.root, text="Alati i postavke")
        fs.pack(fill="x", padx=8, pady=(8, 4))
        fs.columnconfigure(1, weight=1)

        self._path_row(fs, 0, "accoreconsole.exe:", self.var_accore, self._browse_accore)
        self._path_row(fs, 1, "LayoutCreatorCore.dll:", self.var_dll, self._browse_dll)
        self._path_row(fs, 2, "sastAu.dwg:", self.var_sast, self._browse_sast)

        opts = ttk.Frame(fs)
        opts.grid(row=3, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(opts, text="timeout (s):").pack(side="left")
        ttk.Spinbox(opts, from_=30, to=3600, increment=30, width=7,
                    textvariable=self.var_timeout).pack(side="left", padx=(3, 14))
        ttk.Label(opts, text="jobs:").pack(side="left")
        ttk.Spinbox(opts, from_=1, to=8, width=4,
                    textvariable=self.var_jobs).pack(side="left", padx=(3, 14))
        ttk.Checkbutton(opts, text="backup (.bak)", variable=self.var_backup).pack(side="left", padx=(0, 14))
        ttk.Button(opts, text="Učitaj config…", command=self._load_config_dialog).pack(side="left", padx=3)
        ttk.Button(opts, text="Spremi config…", command=self._save_config_dialog).pack(side="left", padx=3)

        # Koraci (plan V2) + Excel
        fk = ttk.LabelFrame(self.root, text="Koraci")
        fk.pack(fill="x", padx=8, pady=4)
        fk.columnconfigure(1, weight=1)

        kbar = ttk.Frame(fk)
        kbar.grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=3)
        ttk.Checkbutton(kbar, text="layouti", variable=self.var_k_layouti).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(kbar, text="polja (Excel→DWG)", variable=self.var_k_polja).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(kbar, text="naslovi (Excel→sastAu)", variable=self.var_k_naslovi).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(kbar, text="sortiranje tabova", variable=self.var_k_sortiranje).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(kbar, text="export (DWG→Excel)", variable=self.var_k_export).pack(side="left")

        self._path_row(fk, 1, "Excel datoteka:", self.var_excel, self._browse_excel)
        sheets = ttk.Frame(fk)
        sheets.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
        ttk.Label(sheets, text="sheet polja:").pack(side="left")
        ttk.Entry(sheets, textvariable=self.var_sheet_polja, width=14).pack(side="left", padx=(3, 14))
        ttk.Label(sheets, text="sheet nacrti:").pack(side="left")
        ttk.Entry(sheets, textvariable=self.var_sheet_nacrti, width=14).pack(side="left", padx=3)

        # DWG projekti
        fd = ttk.LabelFrame(self.root, text="DWG projekti")
        fd.pack(fill="both", expand=True, padx=8, pady=4)

        bar = ttk.Frame(fd)
        bar.pack(fill="x", **pad)
        ttk.Button(bar, text="Dodaj DWG…", command=self._add_dwgs).pack(side="left", padx=3)
        ttk.Button(bar, text="Dodaj folder…", command=self._add_folder).pack(side="left", padx=3)
        ttk.Label(bar, text="pattern:").pack(side="left", padx=(10, 2))
        ttk.Entry(bar, textvariable=self.var_pattern, width=16).pack(side="left")
        ttk.Button(bar, text="Ukloni", command=self._remove_selected).pack(side="right", padx=3)
        ttk.Button(bar, text="Očisti", command=self._clear_files).pack(side="right", padx=3)

        cols = ("status", "excel")
        self.tree = ttk.Treeview(fd, columns=cols, show="tree headings", height=8, selectmode="extended")
        self.tree.heading("#0", text="datoteka")
        self.tree.heading("status", text="status")
        self.tree.heading("excel", text="excel (po projektu)")
        self.tree.column("#0", width=440)
        self.tree.column("status", width=100, anchor="center")
        self.tree.column("excel", width=180)
        self.tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # Akcije
        fa = ttk.Frame(self.root)
        fa.pack(fill="x", padx=8, pady=2)
        self.btn_run = ttk.Button(fa, text="▶  Pokreni", command=lambda: self._start(dry_run=False))
        self.btn_run.pack(side="left", padx=3)
        self.btn_dry = ttk.Button(fa, text="Dry-run", command=lambda: self._start(dry_run=True))
        self.btn_dry.pack(side="left", padx=3)
        self.btn_stop = ttk.Button(fa, text="■  Prekini", command=self._request_abort, state="disabled")
        self.btn_stop.pack(side="left", padx=3)
        self.btn_csv = ttk.Button(fa, text="Otvori CSV", command=self._open_csv, state="disabled")
        self.btn_csv.pack(side="right", padx=3)
        self.btn_openlog = ttk.Button(fa, text="Otvori log", command=self._open_log, state="disabled")
        self.btn_openlog.pack(side="right", padx=3)

        # Napredak
        fp = ttk.Frame(self.root)
        fp.pack(fill="x", padx=8, pady=(2, 0))
        self.progress = ttk.Progressbar(fp, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.lbl_prog = ttk.Label(fp, text="0/0 datoteka", width=22, anchor="e")
        self.lbl_prog.pack(side="right", padx=(6, 0))

        # Log pane (read-only)
        self.txt = tk.Text(self.root, height=12, wrap="none", state="disabled",
                           font=("Consolas", 9), background="#0C1A28", foreground="#D6E6F2")
        self.txt.pack(fill="both", expand=True, padx=8, pady=6)
        sb = ttk.Scrollbar(self.txt, command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.config(yscrollcommand=sb.set)

        self.status = ttk.Label(self.root, text="Spreman.", anchor="w", relief="sunken")
        self.status.pack(fill="x", side="bottom")

    def _path_row(self, parent, r, label, var, browse) -> None:
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(parent, text="…", width=3, command=browse).grid(row=r, column=2, padx=6, pady=3)

    # ---- log helper ----
    def _log(self, text: str) -> None:
        self.txt.config(state="normal")
        self.txt.insert("end", text + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")

    # ---- odabir datoteka ----
    def _add_dwgs(self) -> None:
        paths = filedialog.askopenfilenames(title="Odaberi DWG datoteke",
                                            filetypes=[("AutoCAD DWG", "*.dwg"), ("Sve", "*.*")])
        self._add_paths(Path(p) for p in paths)

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Odaberi folder s DWG projektima")
        if not folder:
            return
        pattern = self.var_pattern.get().strip() or DEFAULT_PATTERN
        found = [p for p in sorted(Path(folder).glob(pattern)) if p.suffix.lower() == ".dwg"]
        if not found:
            messagebox.showinfo("Nema DWG-ova", f"U '{folder}' nema datoteka za pattern '{pattern}'.")
            return
        self._add_paths(found)

    def _add_paths(self, paths) -> None:
        existing = {str(pr.dwg) for pr in self.files}
        for p in paths:
            if p.suffix.lower() == ".dwg" and str(p) not in existing:
                self.files.append(core.Project(p))
                existing.add(str(p))
        self._refresh_tree()

    def _remove_selected(self) -> None:
        sel = set(self.tree.selection())
        self.files = [pr for iid, pr in zip(self._iids(), self.files) if iid not in sel]
        self._refresh_tree()

    def _clear_files(self) -> None:
        self.files.clear()
        self._refresh_tree()

    def _iids(self) -> list[str]:
        return list(self.tree.get_children())

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for pr in self.files:
            self.tree.insert("", "end", text=str(pr.dwg),
                             values=(file_status(pr.dwg),
                                     str(pr.excel) if pr.excel else ""))

    # ---- browse ----
    def _browse_accore(self) -> None:
        p = filedialog.askopenfilename(title="accoreconsole.exe",
                                       filetypes=[("exe", "*.exe"), ("Sve", "*.*")])
        if p:
            self.var_accore.set(p)

    def _browse_dll(self) -> None:
        p = filedialog.askopenfilename(title="LayoutCreatorCore.dll",
                                       filetypes=[("dll", "*.dll"), ("Sve", "*.*")])
        if p:
            self.var_dll.set(p)

    def _browse_sast(self) -> None:
        p = filedialog.askopenfilename(title="sastAu.dwg",
                                       filetypes=[("AutoCAD DWG", "*.dwg"), ("Sve", "*.*")])
        if p:
            self.var_sast.set(p)

    def _browse_excel(self) -> None:
        p = filedialog.askopenfilename(title="Excel datoteka",
                                       filetypes=[("Excel", "*.xlsx *.xlsm"), ("Sve", "*.*")])
        if p:
            self.var_excel.set(p)

    # ---- config load/save (isti format kao CLI) ----
    def _load_initial_config(self) -> None:
        path = None
        try:
            if STATE_FILE.is_file():
                import json
                last = json.loads(STATE_FILE.read_text(encoding="utf-8")).get("last_config")
                if last and Path(last).is_file():
                    path = Path(last)
        except Exception:
            path = None
        if path is None and DEFAULT_CONFIG.is_file():
            path = DEFAULT_CONFIG
        if path is not None:
            self._load_config(path, announce=False)

    def _remember_config(self, path: Path) -> None:
        try:
            import json
            STATE_FILE.write_text(json.dumps({"last_config": str(path)}), encoding="utf-8")
        except Exception:
            pass

    def _load_config_dialog(self) -> None:
        p = filedialog.askopenfilename(title="Učitaj config.json",
                                       initialdir=str(core.SCRIPT_DIR),
                                       filetypes=[("JSON", "*.json"), ("Sve", "*.*")])
        if p:
            self._load_config(Path(p), announce=True)

    def _load_config(self, path: Path, announce: bool) -> None:
        try:
            raw = core.read_config_raw(path)
        except Exception as exc:
            messagebox.showerror("Greška", f"Ne mogu učitati config:\n{exc}")
            return
        self.var_accore.set(str(raw.get("accoreconsole", "")))
        self.var_dll.set(str(raw.get("plugin_dll", "")))
        self.var_sast.set(str(raw.get("sastavnica", "")))
        self.var_timeout.set(int(raw.get("timeout_s", 300)))
        self.var_jobs.set(int(raw.get("jobs", 1)))
        self.var_backup.set(bool(raw.get("backup", True)))
        self.var_lang = str(raw.get("lang", "en-US"))
        if raw.get("pattern"):
            self.var_pattern.set(str(raw["pattern"]))

        # koraci + excel (plan V2); stari config bez ovih polja -> defaulti
        steps = core.parse_steps(raw.get("koraci"))
        self.var_k_layouti.set(steps.layouti)
        self.var_k_polja.set(steps.polja)
        self.var_k_naslovi.set(steps.naslovi)
        self.var_k_sortiranje.set(steps.sortiranje)
        self.var_k_export.set(steps.export)
        self.var_excel.set(str(raw.get("excel", "") or ""))
        self.var_sheet_polja.set(str(raw.get("sheet_polja", "Podaci")))
        self.var_sheet_nacrti.set(str(raw.get("sheet_nacrti", "Nacrti")))

        # projekti / folder+pattern — isti parser kao CLI (string ILI {dwg, excel})
        self.files = core.resolve_projects(raw, path) if (raw.get("projekti") or raw.get("folder")) else []
        self._refresh_tree()
        self._remember_config(path)
        if announce:
            self._log(f"Učitan config: {path}  ({len(self.files)} projekata)")
        self.status.config(text=f"Config: {path}")

    def _steps(self) -> "core.Steps":
        return core.Steps(
            layouti=bool(self.var_k_layouti.get()),
            polja=bool(self.var_k_polja.get()),
            naslovi=bool(self.var_k_naslovi.get()),
            sortiranje=bool(self.var_k_sortiranje.get()),
            export=bool(self.var_k_export.get()),
        )

    def _config_dict(self) -> dict:
        s = self._steps()
        d = {
            "accoreconsole": self.var_accore.get(),
            "plugin_dll": self.var_dll.get(),
            "sastavnica": self.var_sast.get(),
            "timeout_s": int(self.var_timeout.get()),
            "backup": bool(self.var_backup.get()),
            "jobs": int(self.var_jobs.get()),
            "lang": self.var_lang,
            "koraci": {"layouti": s.layouti, "polja": s.polja, "naslovi": s.naslovi,
                       "sortiranje": s.sortiranje, "export": s.export},
            "sheet_polja": self.var_sheet_polja.get() or "Podaci",
            "sheet_nacrti": self.var_sheet_nacrti.get() or "Nacrti",
            # string za projekte bez vlastitog excela, {dwg, excel} za ostale
            "projekti": [str(pr.dwg) if pr.excel is None
                         else {"dwg": str(pr.dwg), "excel": str(pr.excel)}
                         for pr in self.files],
        }
        if self.var_excel.get().strip():
            d["excel"] = self.var_excel.get().strip()
        return d

    def _save_config_dialog(self) -> None:
        p = filedialog.asksaveasfilename(title="Spremi config.json",
                                         initialdir=str(core.SCRIPT_DIR),
                                         initialfile="config.json", defaultextension=".json",
                                         filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            core.write_config(Path(p), self._config_dict())
            self._remember_config(Path(p))
            self._log(f"Spremljen config: {p}")
        except Exception as exc:
            messagebox.showerror("Greška", f"Ne mogu spremiti config:\n{exc}")

    # ---- pokretanje ----
    def _build_config(self) -> core.Config | None:
        steps = self._steps()
        required = [("accoreconsole.exe", self.var_accore.get()),
                    ("LayoutCreatorCore.dll", self.var_dll.get())]
        if steps.layouti:
            # sastavnica je potrebna samo za korak 'layouti' (kao CLI preflight)
            required.append(("sastAu.dwg", self.var_sast.get()))
        for label, val in required:
            if not val.strip():
                messagebox.showwarning("Nedostaje putanja", f"Postavi: {label}")
                return None
            if not Path(val).is_file():
                if not messagebox.askyesno("Putanja ne postoji",
                                           f"{label} ne postoji:\n{val}\n\nSvejedno nastavi?"):
                    return None
        if not any((steps.layouti, steps.polja, steps.naslovi, steps.sortiranje, steps.export)):
            messagebox.showwarning("Nema koraka", "Uključi barem jedan korak.")
            return None
        excel = self.var_excel.get().strip()
        if (steps.polja or steps.naslovi or steps.export) and not excel \
                and not any(pr.excel for pr in self.files):
            messagebox.showwarning(
                "Nedostaje Excel",
                "Koraci polja/naslovi/export trebaju Excel datoteku\n"
                "(globalnu ili po projektu).")
            return None
        if not self.files:
            messagebox.showwarning("Nema projekata", "Dodaj barem jedan DWG.")
            return None
        return core.Config(
            accoreconsole=Path(self.var_accore.get()),
            plugin_dll=Path(self.var_dll.get()),
            sastavnica=Path(self.var_sast.get()) if self.var_sast.get().strip() else Path("sastAu.dwg"),
            timeout_s=int(self.var_timeout.get()),
            backup=bool(self.var_backup.get()),
            jobs=int(self.var_jobs.get()),
            lang=self.var_lang,
            excel=Path(excel) if excel else None,
            sheet_polja=self.var_sheet_polja.get() or "Podaci",
            sheet_nacrti=self.var_sheet_nacrti.get() or "Nacrti",
            koraci=steps,
            projekti=list(self.files),
        )

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        for b in (self.btn_run, self.btn_dry, self.btn_csv, self.btn_openlog):
            b.config(state=state)
        self.btn_stop.config(state="normal" if running else "disabled")

    def _start(self, dry_run: bool) -> None:
        if self.worker and self.worker.is_alive():
            return
        cfg = self._build_config()
        if cfg is None:
            return
        self._refresh_tree()

        if dry_run:
            self._log("\n--- DRY RUN (ništa se ne mijenja) ---")
            self._log(f"Koraci: {core._steps_summary(cfg.koraci)}")
            self._log("Create-pass naredbe:")
            for c in core.build_commands(cfg.koraci):
                self._log(f"  {c if c else '(prazna linija = kraj selekcije)'}")
            for pr in cfg.projekti:
                xls = pr.excel or cfg.excel
                xnote = f"  excel={xls}" if (cfg.koraci.polja or cfg.koraci.naslovi
                                             or cfg.koraci.export) else ""
                self._log(f"  {file_status(pr.dwg):<10}  {pr.dwg}{xnote}")
            self._log("--- kraj dry-run ---")
            return

        self.abort.clear()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log = core.setup_logging(ts, to_stdout=False)
        log.addHandler(QueueLogHandler(self.q))
        self.last_csv = None
        self.last_log = core.run_log_path(ts)

        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")
        self.progress.config(value=0, maximum=len(cfg.projekti))
        self.lbl_prog.config(text=f"0/{len(cfg.projekti)} datoteka")
        self._set_running(True)
        self.status.config(text="Obrada u tijeku…")

        self.worker = threading.Thread(target=self._run_worker, args=(cfg, log, ts), daemon=True)
        self.worker.start()

    def _request_abort(self) -> None:
        self.abort.set()
        self._log(">>> PREKID zatražen — dovršavam/ubijam tekuću datoteku, ostale preskačem…")
        self.btn_stop.config(state="disabled")

    # ---- worker thread (NE dira tkinter; sve ide kroz self.q) ----
    def _run_worker(self, cfg: core.Config, log: logging.Logger, ts: str) -> None:
        all_rows: list[core.Row] = []
        total = len(cfg.projekti)
        try:
            for i, pr in enumerate(cfg.projekti, 1):
                name = pr.dwg.name
                if self.abort.is_set():
                    self.q.put(("log", f"[{i}/{total}] preskačem (prekid): {name}"))
                    all_rows.append(core.Row(name, "", "", core.ERR, "preskoceno (prekid)"))
                    continue
                self.q.put(("log", f"[{i}/{total}] {name} — obrada…"))

                def on_progress(done, tot, layout, _n=name):
                    self.q.put(("prog_layout", _n, done, tot, layout))

                rows = core.process_file(cfg, pr, log,
                                         on_progress=on_progress,
                                         should_abort=self.abort.is_set)
                all_rows.extend(rows)
                self.q.put(("prog_file", i, total))

            csv_path = core.write_csv(ts, all_rows)
            n_ok = sum(1 for r in all_rows if r.status == core.OK)
            n_warn = sum(1 for r in all_rows if r.status == core.WARN)
            n_err = sum(1 for r in all_rows if r.status == core.ERR)
            self.q.put(("summary", n_ok, n_warn, n_err, str(csv_path)))
        except Exception as exc:  # noqa: BLE001
            self.q.put(("log", f"NEOČEKIVANA GREŠKA: {exc}"))
            self.q.put(("summary", 0, 0, -1, ""))
        finally:
            self.q.put(("finished",))

    # ---- glavni thread prazni queue ----
    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._log(msg[1])
                elif kind == "prog_layout":
                    _n, done, tot, layout = msg[1], msg[2], msg[3], msg[4]
                    t = f"/{tot}" if tot else ""
                    self.status.config(text=f"{_n}: {done}{t} layouta  ({layout})")
                elif kind == "prog_file":
                    i, total = msg[1], msg[2]
                    self.progress.config(value=i, maximum=total)
                    self.lbl_prog.config(text=f"{i}/{total} datoteka")
                elif kind == "summary":
                    n_ok, n_warn, n_err, csv_path = msg[1], msg[2], msg[3], msg[4]
                    self._log(f"\n=== SAŽETAK === OK={n_ok}  WARN={n_warn}  ERR={n_err}")
                    if csv_path:
                        self.last_csv = Path(csv_path)
                        self._log(f"CSV: {csv_path}")
                    if self.last_log:
                        self._log(f"Log: {self.last_log}")
                    self.status.config(text=f"Gotovo: OK={n_ok} WARN={n_warn} ERR={n_err}")
                elif kind == "finished":
                    self._set_running(False)
                    self.btn_csv.config(state="normal" if self.last_csv else "disabled")
                    self.btn_openlog.config(state="normal" if self.last_log else "disabled")
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    # ---- otvaranje rezultata ----
    def _open_path(self, p: Path | None) -> None:
        if not p or not p.exists():
            messagebox.showinfo("Nema datoteke", "Datoteka još ne postoji.")
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(p))  # type: ignore[attr-defined]  # Windows
            else:
                messagebox.showinfo("Putanja", str(p))
        except Exception as exc:
            messagebox.showerror("Greška", str(exc))

    def _open_csv(self) -> None:
        self._open_path(self.last_csv)

    def _open_log(self) -> None:
        self._open_path(self.last_log)


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    BatchGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
