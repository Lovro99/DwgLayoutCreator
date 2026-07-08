"""
PoC — dokaz kritičnog koraka za batch layout pristup:
accoreconsole.exe (headless AutoCAD core) može otvoriti DWG, kreirati layout,
spremiti datoteku, i to se može verificirati drugim headless passom.

NIJE dio produkcijskog alata — samo dokaz izvedivosti prije implementacije
po IMPLEMENTACIJA_PLAN.md.

Korištenje (na Windows stroju s instaliranim AutoCAD-om 2025):
    python poc_accoreconsole.py "C:\\putanja\\do\\projekta.dwg"

Radi ISKLJUČIVO na kopiji u temp direktoriju — original se ne dira.
Ispisuje PASS ili FAIL.
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ACCORE = Path(r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe")
LAYOUT_NAME = "POC_TEST"
TIMEOUT_S = 180

# Pass 1: kreiraj layout i spremi.
# VAŽNO: .scr datoteke drži ASCII-only (bez č/ć/š/đ/ž) — accoreconsole zna
# krivo interpretirati ne-ASCII znakove ovisno o codepageu.
SCR_CREATE = f'(command "._-LAYOUT" "_New" "{LAYOUT_NAME}")\n._QSAVE\n'

# Pass 2: ispisi imena layouta u tekstualnu datoteku — cisti entity-AutoLISP.
# VAZNO (nauceno u Koraku 0): (layoutlist) NE postoji u accoreconsole core
# konzoli ("; error: no function definition: LAYOUTLIST") — ta se funkcija
# definira tek u punom AutoCAD startup LISP-u. Isto vrijedi za ActiveX
# (vla-*/vlax-*). Zato imena citamo iz ACAD_LAYOUT rjecnika preko
# namedobjdict/dictsearch (osnovni AutoLISP, radi u core konzoli).
# U DXF podacima rjecnika svaki unos ima par (3 . "ime layouta").
SCR_VERIFY = (
    '(setq f (open "{out}" "w"))\n'
    '(setq d (dictsearch (namedobjdict) "ACAD_LAYOUT"))\n'
    "(foreach pair d (if (= 3 (car pair)) (write-line (cdr pair) f)))\n"
    "(close f)\n"
)


def run_scr(dwg: Path, scr_text: str, tag: str) -> str:
    scr = dwg.parent / f"poc_{tag}.scr"
    scr.write_text(scr_text, encoding="ascii")
    cmd = [str(ACCORE), "/i", str(dwg), "/s", str(scr), "/l", "en-US"]
    print(f"\n>> [{tag}]", " ".join(cmd))
    res = subprocess.run(cmd, capture_output=True, timeout=TIMEOUT_S)
    out = res.stdout.decode("utf-8", errors="replace")
    print(out[-1500:])
    if res.returncode != 0:
        sys.exit(f"FAIL — accoreconsole exit code {res.returncode} u passu '{tag}'")
    return out


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    if not ACCORE.exists():
        sys.exit(f"FAIL — nema {ACCORE} (prilagodi putanju za svoju verziju AutoCAD-a)")
    src = Path(sys.argv[1])
    if not src.is_file():
        sys.exit(f"FAIL — ne postoji: {src}")

    work = Path(tempfile.mkdtemp(prefix="poc_layout_")) / src.name
    shutil.copy2(src, work)
    print(f"Radim na kopiji: {work}")

    run_scr(work, SCR_CREATE, "create")

    out_txt = work.parent / "layouts.txt"
    # LISP open() trazi forward slasheve
    run_scr(work, SCR_VERIFY.format(out=str(out_txt).replace("\\", "/")), "verify")

    if not out_txt.exists():
        sys.exit("FAIL — verifikacijski pass nije zapisao layouts.txt")
    layouts = out_txt.read_text(encoding="utf-8", errors="replace").split()
    print("\nLayouti nakon round-tripa:", layouts)

    if LAYOUT_NAME in layouts:
        print(f"\nPASS — accoreconsole je headless kreirao i spremio layout '{LAYOUT_NAME}'.")
    else:
        sys.exit(f"FAIL — layout '{LAYOUT_NAME}' nije pronadjen nakon spremanja.")


if __name__ == "__main__":
    main()
