# DwgLayoutCreator

Alati za automatsku obradu paperspace layouta u AutoCAD `.dwg` projektima —
headless, bez otvaranja AutoCAD GUI-a i bez Excela:

- **kreiranje layouta** (viewport + sastavnica + page setup) iz blokova
  `A4L…A0P`/`CUSTOM1-8` u model spaceu,
- **polja** — upis custom drawing properties iz Excela (port `SetFieldsValue.lsp`),
- **naslovi** — upis naslova/mjerila u `sastAu` blok po layoutu (port `SetLayoutTitles.lsp`),
- **sortiranje tabova** — numerički redoslijed layouta preko `TabOrder` (komparator iz `ExportLayoutsToExcel.lsp`),
- **export** — sortirana imena layouta u stupac A Excela (port `ExportLayoutsToExcel.lsp`).

Tri komponente, po zrelosti:

| Komponenta | Što je | Status |
|---|---|---|
| **`batch/` + `LayoutCrator/LayoutCreatorCore`** | Headless batch preko `accoreconsole.exe` (AutoCAD 2025 core) | ✅ **preporučeno** |
| `LayoutCrator/LayoutCreator` | AutoCAD plugin (NETLOAD), interaktivna komanda `CREATELAYOUT` | ✅ radi u AutoCAD-u |
| `DwgLayoutCreator/` (standalone, ACadSharp) | Pokušaj "uredi DWG bez AutoCAD-a" | ❄️ **zamrznuto** (slijepa ulica — vidi `IMPLEMENTACIJA_PLAN.md §1.2`); ostaviti samo kao read-only referencu |

Detaljna dijagnoza i obrazloženje odabira: **`IMPLEMENTACIJA_PLAN.md`**;
specifikacija 4 dodatna koraka (polja/naslovi/sortiranje/export):
**`IMPLEMENTACIJA_PLAN_V2.md`**.

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
  "koraci": { "layouti": true, "polja": false, "naslovi": false,
              "sortiranje": false, "export": false },
  "excel":         "C:\\Projekti\\zajednicki_podaci.xlsx",
  "sheet_polja":   "Podaci",
  "sheet_nacrti":  "Nacrti",
  "projekti": [
    "C:\\Projekti\\P100\\elektro.dwg",
    { "dwg": "C:\\Projekti\\P101\\elektro.dwg", "excel": "C:\\Projekti\\P101\\podaci.xlsx" }
  ]
}
```

| Polje | Značenje |
|---|---|
| `accoreconsole` | Putanja do `accoreconsole.exe`. |
| `plugin_dll` | Putanja do buildanog `LayoutCreatorCore.dll`. |
| `sastavnica` | `sastAu.dwg` — šalje se pluginu kroz env var `LAYOUT_SAST_PATH` (potrebna samo za korak `layouti`). |
| `timeout_s` | Max sekundi po datoteci; prekoračenje → kill → `ERR` (original netaknut). |
| `backup` | `true` → original ide u `<ime>.dwg.bak` prije zamjene. |
| `jobs` | Paralelnih `accoreconsole` procesa (>1 gasi live progress zbog ispisa). |
| `koraci` | Koji se koraci izvršavaju (v. tablicu ispod). Izostavljeno = samo `layouti` (staro ponašanje — stari config radi identično). |
| `excel` | Globalna Excel datoteka za korake polja/naslovi/export (`.xlsx`/`.xlsm`; stari `.xls` NIJE podržan). Po projektu se može nadjačati (`{dwg, excel}` zapis). |
| `sheet_polja` | Sheet za korak `polja` (default `Podaci`): stupac A = ime svojstva, B = vrijednost. |
| `sheet_nacrti` | Sheet za korake `naslovi`/`export` (default `Nacrti`): A = ime layouta, B = naslov, C = mjerilo. |
| `projekti` | Lista punih putanja DWG-ova (string ILI `{"dwg": "...", "excel": "..."}`)… |
| `folder` + `pattern` | …ILI umjesto `projekti`: `"folder": "C:\\Projekti"`, `"pattern": "**/elektro*.dwg"`. |

Sve o samim layoutima (format papira, broj, položaj) čita se iz **blokova u DWG-u**
(`A4L…A0P`/`CUSTOM1-8` + atribut broja `BR_N` ili zadnji atribut) — config to ne duplicira.

### Koraci po datoteci (svi u ISTOM prolazu, redom)

| Korak | Komanda / izvršitelj | Port LISP-a | Što radi |
|---|---|---|---|
| `layouti` | `CREATELAYOUTBATCH` | `createlayoutV2.lsp` | Kreira layoute iz blokova u model spaceu (kao dosad). |
| `polja` | `SETFIELDSBATCH` + `UPDATEFIELD _All` | `SetFieldsValue.lsp` | Custom drawing properties iz Excela (`Podaci`): postojeći ključ prepiši, novi dodaj; pa osvježi FIELD-ove. |
| `naslovi` | `SETTITLESBATCH` | `SetLayoutTitles.lsp` | Za svaki red `Nacrti` sheeta upiše `layout_title`/`layout_mjerilo` u prvi `sastAu` blok tog layouta. |
| `sortiranje` | `SORTTABSBATCH` | `TabSortV2-2` jezgra + `ele:layout<` | Tabovi imena `broj-brojxbroj` numerički (prvi pa drugi broj); ostali iza njih. |
| `export` | Python (nakon verifyja) | `ExportLayoutsToExcel.lsp` | Sortirana valjana imena u stupac A `Nacrti` sheeta (retci ispod N se NE brišu — kao LISP). |

Excel se čita/piše **bez Excela** (`batch/xlsx_lite.py`, stdlib ZIP+XML); podaci
pluginu idu kroz temp JSON + env var `LAYOUT_BATCH_JSON`. Verifikacija:
`VERIFYBATCH` (read-only) dumpa layoute+`TabOrder`, custom propertyje i
naslove; orkestrator uspoređuje sa poslanim — mismatch → `ERR`, original netaknut.

Tipičan tok kroz vrijeme: 1. batch s `layouti`+`sortiranje`+`export` → 2. u
Excelu (`Nacrti`) upišeš naslove/mjerila uz izvezena imena → 3. batch s
`naslovi` (+`polja`) → naslovi u sastavnicama. Svi koraci su idempotentni.

### 3. Pokretanje — GUI ili CLI

**GUI** (bez ručnog editiranja configa) — `python batch\batch_gui.py` (ili launcher
`Batch Layouti.bat`): odabir DWG-ova (gumbi / folder + pattern) sa statusom
(postoji/zaključan/read-only), Browse za alate, Učitaj/Spremi **isti** `config.json`,
gumbi Pokreni / Dry-run / Prekini, live log + progress bar + Otvori CSV/log.
GUI i CLI dijele isti config format.

**CLI**:
```
python batch\batch_layouts.py --config batch\config.json --dry-run   # samo ispiši što bi radio
python batch\batch_layouts.py --config batch\config.json             # stvarno
python batch\batch_layouts.py --config batch\config.json --jobs 3    # paralelno
```
Po datoteci: kopija → **create pass** (`NETLOAD` → uključeni koraci → `QSAVE`)
→ **verify pass** (`VERIFYBATCH`, read-only) → usporedbe (layouti / polja /
naslovi / poredak tabova) → tek onda original → `.bak`, kopija na mjesto
originala → export u Excel.
**Ako bilo što pukne (lock, timeout, sastavnica fali, verify FAIL) → original ostaje netaknut.**

Rezultati: `batch/logs/rezultati_<timestamp>.csv` (UTF-8 BOM, `;`-delimited,
stupci `datoteka;korak;layout;status;poruka`) + `batch/logs/run_<timestamp>.log`.
Idempotentno — ponovno pokretanje preskače postojeće layoute (`created=0`), a
polja/naslovi/sortiranje/export samo ponovno upišu iste vrijednosti.

U GUI-ju su koraci checkboxovi (+ polja za Excel i imena sheetova); GUI i CLI
i dalje dijele isti `config.json`.

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
| `UPDATEFIELD` javlja *Unknown command* u logu | Core konzola bez te komande — bezopasno (WARN): FIELD-ovi se ionako reevaluiraju pri otvaranju/plotu (REGEN). Javi ako se pojavi, da se doda REGEN fallback u `.scr`. |
| `xls nije podrzan` | Stari binarni `.xls` — spremi kao `.xlsx`/`.xlsm`. |
| `export u Excel nije uspio` (PermissionError) | Excel datoteka otvorena u Excelu — zatvori ju pa ponovi samo `export` korak (DWG izmjene su već valjane). |
| `TITLES\|noblock\|…` | Layout nema `sastAu` blok ili blok nema tagove `layout_title`/`layout_mjerilo`. |
| `TITLES\|nolayout\|…` | Ime layouta u `Nacrti` sheetu ne postoji u DWG-u (tipfeler ili layout još nije kreiran). |
| **polja se ne upišu** (`verify FAIL: polja: kljuc … nije upisan`) | U `run_*.log` je sad i *create stdout*: pogledaj `RESULT\|fields\|added=… present=P/Q` i eventualne `ERR\| fields:` linije. `present<Q` = upis nije "sjeo" ni u sesiji (iznimka na setteru — vidi `ERR\|`); `present=Q` a `VPROP` prazan = problem spremanja (javi log). |

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
├── IMPLEMENTACIJA_PLAN.md          # dijagnoza, odabir pristupa, koraci (V1)
├── IMPLEMENTACIJA_PLAN_V2.md       # specifikacija koraka polja/naslovi/sortiranje/export
├── batch/                          # Python orkestrator (headless batch)
│   ├── batch_layouts.py            # CLI + logika obrade (importabilan modul)
│   ├── batch_gui.py                # tkinter GUI (dijeli config s CLI-jem)
│   ├── layout_names.py             # port ele:nums/ele:layout< (filter + komparator)
│   ├── xlsx_lite.py                # čitanje/pisanje .xlsx bez Excela (stdlib ZIP+XML)
│   ├── tests/                      # unit testovi (python -m unittest discover -s batch/tests)
│   ├── config.json                 # primjer
│   ├── templates/
│   │   ├── create_layouts.scr.tpl  # NETLOAD + __COMMANDS__ (dinamički koraci)
│   │   └── verify_layouts.scr.tpl  # NETLOAD + VERIFYBATCH
│   └── logs/                       # rezultati_*.csv + run_*.log (gitignore)
├── LayoutCrator/
│   ├── LayoutCreator/              # AutoCAD plugin (NETLOAD) — dijeljeni .cs
│   └── LayoutCreatorCore/          # accoreconsole (core) build
│       ├── LayoutCreatorCore.csproj
│       ├── CreateLayoutBatchCommand.cs   # CREATELAYOUTBATCH
│       ├── SetFieldsBatchCommand.cs      # SETFIELDSBATCH  (SetFieldsValue.lsp port)
│       ├── SetTitlesBatchCommand.cs      # SETTITLESBATCH  (SetLayoutTitles.lsp port)
│       ├── SortTabsBatchCommand.cs       # SORTTABSBATCH   (TabSort jezgra)
│       ├── VerifyBatchCommand.cs         # VERIFYBATCH     (read-only verifikacija)
│       ├── LayoutNameOrder.cs            # ele:nums/ele:layout< port (identican layout_names.py)
│       ├── BatchDataFile.cs              # LAYOUT_BATCH_JSON čitanje
│       └── netload_diag.scr.tpl
├── poc/
│   └── poc_accoreconsole.py        # PoC headless lanca (Korak 0)
└── DwgLayoutCreator/               # ❄️ standalone (ACadSharp) — zamrznuto, read-only referenca
```

## Ručna provjera (zahtijeva AutoCAD — nije izvedivo u sandboxu)

1. `dotnet build LayoutCrator\LayoutCreatorCore -c Release` — mora proći.
2. Batch s `koraci: {layouti: true, sortiranje: true}` na kopiji test projekta →
   tabovi numerički poredani, CSV bez `ERR`.
3. Excel s `Podaci` (2-3 para) i `Nacrti` (naslov+mjerilo za 2 layouta) → batch s
   `polja`+`naslovi` → u AutoCAD-u provjeri DWGPROPS custom properties, naslove u
   sastavnici, i da FIELD-ovi pokazuju nove vrijednosti.
4. Batch s `export: true` → stupac A `Nacrti` sheeta == sortirana imena layouta.
5. Idempotentnost: isti batch 2× → drugi prolaz 0 promjena, 0 `ERR`.

## Ovisnosti
- AutoCAD 2025 (`accoreconsole.exe`), .NET 8 (x64), Python ≥ 3.10 (stdlib).
- ACadSharp 3.5.7 ostaje SAMO u read-only `dwg_props_helper` (EE-Python-Tools), ne za pisanje DWG-a.
