# DwgLayoutCreator

Alati za automatsko kreiranje paperspace layouta (viewport + sastavnica + page
setup) u AutoCAD `.dwg` projektima na temelju blokova formata `A4L…A0P` i
`CUSTOM1-8` u model spaceu.

Tri komponente, po zrelosti:

| Komponenta | Što je | Status |
|---|---|---|
| **`batch/` + `LayoutCrator/LayoutCreatorCore`** | Headless batch preko `accoreconsole.exe` (AutoCAD 2025 core) | ✅ **preporučeno** |
| `LayoutCrator/LayoutCreator` | AutoCAD plugin (NETLOAD), interaktivna komanda `CREATELAYOUT` | ✅ radi u AutoCAD-u |
| `DwgLayoutCreator/` (standalone, ACadSharp) | Pokušaj "uredi DWG bez AutoCAD-a" | ❄️ **zamrznuto** (slijepa ulica — vidi `IMPLEMENTACIJA_PLAN.md §1.2`); ostaviti samo kao read-only referencu |

Detaljna dijagnoza i obrazloženje odabira: **`IMPLEMENTACIJA_PLAN.md`**.

---

## Batch layouti (headless) — preporučeni tok

Lanac: **Python orkestrator → `accoreconsole.exe` → NETLOAD `LayoutCreatorCore.dll` → `CREATELAYOUTBATCH`**.
Ponovno koristi istu, već izdebugiranu domensku logiku kao interaktivni plugin
(scanner dinamičkih blokova, `PlotSettingsValidator` page setup, sastavnica,
pravi `GeometricExtents`), samo headless i batch.

### Preduvjeti
- **AutoCAD 2025** (puni — `accoreconsole.exe` NE dolazi uz LT).
- **.NET SDK 8**.
- **Python ≥ 3.10** (samo standardna biblioteka — bez third-party paketa).
- Kopija sastavnice **`sastAu.dwg`**.

### 1. Build core plugina
```
cd LayoutCrator\LayoutCreatorCore
dotnet build -c Release
```
Rezultat: `bin\Release\net8.0-windows\LayoutCreatorCore.dll`.
*(3 × `MSB3277` upozorenja su bezopasna — AutoCAD DLL-ovi vs .NET 8 ref pack; utišana.)*

### 2. Konfiguracija — `batch/config.json`
```json
{
  "accoreconsole": "C:\\Program Files\\Autodesk\\AutoCAD 2025\\accoreconsole.exe",
  "plugin_dll":    "C:\\...\\LayoutCreatorCore\\bin\\Release\\net8.0-windows\\LayoutCreatorCore.dll",
  "sastavnica":    "C:\\...\\sastAu.dwg",
  "timeout_s":     300,
  "backup":        true,
  "jobs":          1,
  "lang":          "en-US",
  "projekti": [
    "C:\\Projekti\\P100\\elektro.dwg",
    "C:\\Projekti\\P101\\elektro.dwg"
  ]
}
```

| Polje | Značenje |
|---|---|
| `accoreconsole` | Putanja do `accoreconsole.exe`. |
| `plugin_dll` | Putanja do buildanog `LayoutCreatorCore.dll`. |
| `sastavnica` | `sastAu.dwg` — šalje se pluginu kroz env var `LAYOUT_SAST_PATH`. |
| `timeout_s` | Max sekundi po datoteci; prekoračenje → kill → `ERR` (original netaknut). |
| `backup` | `true` → original ide u `<ime>.dwg.bak` prije zamjene. |
| `jobs` | Paralelnih `accoreconsole` procesa (>1 gasi live progress zbog ispisa). |
| `projekti` | Lista punih putanja DWG-ova… |
| `folder` + `pattern` | …ILI umjesto `projekti`: `"folder": "C:\\Projekti"`, `"pattern": "**/elektro*.dwg"`. |

Sve o samim layoutima (format papira, broj, položaj) čita se iz **blokova u DWG-u**
(`A4L…A0P`/`CUSTOM1-8` + atribut broja `BR_N` ili zadnji atribut) — config to ne duplicira.

### 3. Pokretanje
```
python batch\batch_layouts.py --config batch\config.json --dry-run   # samo ispiši što bi radio
python batch\batch_layouts.py --config batch\config.json             # stvarno
python batch\batch_layouts.py --config batch\config.json --jobs 3    # paralelno
```
Po datoteci: kopija → **create pass** (`NETLOAD` → `CREATELAYOUTBATCH` → `QSAVE`)
→ **verify pass** (čita `ACAD_LAYOUT` rječnik) → usporedba da je svako kreirano
ime spremljeno → tek onda original → `.bak`, kopija na mjesto originala.
**Ako bilo što pukne (lock, timeout, sastavnica fali, verify FAIL) → original ostaje netaknut.**

Rezultati: `batch/logs/rezultati_<timestamp>.csv` (UTF-8 BOM, `;`-delimited, jedan
redak po layoutu) + `batch/logs/run_<timestamp>.log`. Idempotentno — ponovno
pokretanje preskače postojeće layoute (`created=0`).

### Dijagnostičke komande (u AutoCAD-u ili headless preko `.scr`)
- `DIAGLAYOUTS` — ispiše svaki blok u model spaceu (efektivno ime, atributi, `dynamic=?`) i koje scanner matcha.
- `LISTMEDIA` — ispiše dostupne plotere i media imena (za troubleshooting page setupa).

---

## Troubleshooting

| Simptom | Uzrok / rješenje |
|---|---|
| **`NETLOAD` blokiran** (`SECURELOAD` / *rejected*) | DLL nije na pouzdanoj putanji. Dodaj folder DLL-a u sysvar `TRUSTEDPATHS` (točka-zarez odvojeno), ili za brzi test na vrh `.scr` stavi `(setvar "SECURELOAD" 0)`. |
| `plugin_dll ne postoji` | Nisi buildao `LayoutCreatorCore` (korak 1) ili je kriva putanja u configu. |
| **timeout > Ns** | Velik crtež — digni `timeout_s` (npr. 900). Ako i dalje pada, u `run_*.log` je zadnjih 2000 znakova stdouta (možda modalni dijalog / educational stamp / proxy grafika). |
| `RESULT\| created>0 bez CREATED\| imena` | Stari DLL — rebuildaj `LayoutCreatorCore`. |
| `zaključana (.dwl/.dwl2)` | DWG otvoren u AutoCAD-u — zatvori ga pa ponovi (batch ga preskače, original netaknut). |
| `sastAu.dwg nedostupan` | Kriva `sastavnica` putanja u configu. |
| Krivi papir / margine | Provjeri `DWG To PDF.pc3` + PMP preko `LISTMEDIA`; `PageSetupConfigurator` radi 4-razinsko matchanje canonical media imena. |

---

## AutoCAD plugin (interaktivno)

`LayoutCrator/LayoutCreator` — NETLOAD u punom AutoCAD-u, pa komanda `CREATELAYOUT`
(+ `DIAGLAYOUTS`, `LISTMEDIA`). Isti izvorni `.cs` dijeli s core buildom;
`CREATELAYOUTBATCH` i core-specifičnosti su u `LayoutCreatorCore` (odabrano preko
`CORECONSOLE` define + `Application` alias). Ne dirati interaktivni tok kod core izmjena.

```
dotnet build LayoutCrator\LayoutCreator\LayoutCreator.csproj -c Release
```

---

## Struktura

```
DwgLayoutCreator/
├── IMPLEMENTACIJA_PLAN.md          # dijagnoza, odabir pristupa, koraci
├── batch/                          # Python orkestrator (headless batch)
│   ├── batch_layouts.py
│   ├── config.json                 # primjer
│   ├── templates/
│   │   ├── create_layouts.scr.tpl
│   │   └── verify_layouts.scr.tpl
│   └── logs/                       # rezultati_*.csv + run_*.log (gitignore)
├── LayoutCrator/
│   ├── LayoutCreator/              # AutoCAD plugin (NETLOAD) — dijeljeni .cs
│   └── LayoutCreatorCore/          # accoreconsole (core) build + CREATELAYOUTBATCH
│       ├── LayoutCreatorCore.csproj
│       ├── CreateLayoutBatchCommand.cs
│       └── netload_diag.scr.tpl
├── poc/
│   └── poc_accoreconsole.py        # PoC headless lanca (Korak 0)
└── DwgLayoutCreator/               # ❄️ standalone (ACadSharp) — zamrznuto, read-only referenca
```

## Ovisnosti
- AutoCAD 2025 (`accoreconsole.exe`), .NET 8 (x64), Python ≥ 3.10 (stdlib).
- ACadSharp 3.5.7 ostaje SAMO u read-only `dwg_props_helper` (EE-Python-Tools), ne za pisanje DWG-a.
