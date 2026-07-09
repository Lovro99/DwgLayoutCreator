# IMPLEMENTACIJA_PLAN_V2.md — Proširenje headless batcha: polja, naslovi, sortiranje tabova, export u Excel

> Plan za implementaciju (Opus). Nadovezuje se na postojeći, RADEĆI pipeline
> (vidi `README.md` i `IMPLEMENTACIJA_PLAN.md`). Grana:
> `claude/fable-layout-automation-rv8gu2` u repoima **DwgLayoutCreator** i
> **AutoLisp**. NAPOMENA o bazi grane: `master` DwgLayoutCreatora sadrži samo
> init commit — pipeline živi na `claude/cad-layout-automation-plan-55yhe2`,
> pa je ova grana bazirana na njoj (svjesno odstupanje od "iz main-a").
> NE otvarati PR.

---

## 0. Cilj

Jedna headless naredba koja nad DWG projektom, u ISTOM prolazu po datoteci,
uz postojeće kreiranje layouta obavi još 4 koraka — bez otvaranja AutoCAD GUI-a
i bez otvaranja Excela:

1. **polja** — upis custom drawing properties iz Excela (port `SetFieldsValue.lsp`)
2. **naslovi** — upis naslova/mjerila u `sastAu` blok po layoutu (port `SetLayoutTitles.lsp`)
3. **sortiranje** — numeričko sortiranje layout tabova preko `TabOrder` (jezgra `TabSortV2-2.lsp` + komparator iz `ExportLayoutsToExcel.lsp`)
4. **export** — ispis sortiranih imena layouta u stupac A Excela (port `ExportLayoutsToExcel.lsp`)

Svi koraci uklopljeni u postojeću garanciju: **kopija → izmjene → verify → tek
onda zamjena originala (uz .bak); ako išta pukne, original NETAKNUT.**

## 1. KLJUČNA ODLUKA — Excel bez Excela (opcija b)

**Odabrano: (b) Python orkestrator sam čita/piše `.xlsx` (stdlib `zipfile` +
`xml.etree`), podatke pluginu prosljeđuje kroz privremeni JSON + env varijablu
(uzor: `LAYOUT_SAST_PATH`).**

Obrazloženje — opcija (a) (reuse `GetExcel.lsp` COM logike kroz accoreconsole)
**nije izvediva**: `GetExcel.lsp` je u cijelosti ActiveX (`vlax-get-or-create-object
"Excel.Application"`), a ActiveX/COM u accoreconsole NE POSTOJI (dokazano u
Koraku 0 v1: čak ni `layoutlist` ni `vla-*` nisu dostupni — vidi komentar u
`poc/poc_accoreconsole.py`). Uz to bi (a) zahtijevala instaliran Excel, lansirala
Excel procese u batchu i naslijedila COM flakiness. Opcija (b): `.xlsx` je ZIP
XML-ova → čitanje i pisanje je čisti stdlib, deterministično, **testabilno i u
ovom (Linux) okruženju bez AutoCAD-a i Excela**, i ne ovisi ni o kakvom vanjskom
softveru. Dosljedno je koristiti kroz SVE korake.

Ograničenja koja se prihvaćaju i dokumentiraju:
- Podržani formati: `.xlsx` i `.xlsm` (oba su ZIP; kod `.xlsm` se `vbaProject.bin`
  ne dira). Stari binarni `.xls` NIJE podržan → ERR s porukom "konvertiraj u .xlsx".
- Pisanje mijenja SAMO ciljani sheet unutar ZIP-a (ostali dijelovi se kopiraju
  bajt-za-bajt); vrijednosti se pišu kao inline stringovi (`t="inlineStr"`) da se
  ne dira `sharedStrings.xml`.

## 2. Arhitektura — što se dodaje gdje

```
batch/batch_layouts.py (orkestrator, SAMO stdlib)
  ├─ čita config.json (proširena shema, §5)
  ├─ čita Excel:  sheet "Podaci" (A=ključ,B=vrijednost) i "Nacrti" (A,B,C)  [xlsx_lite.py]
  ├─ piše temp JSON po DWG-u → env var LAYOUT_BATCH_JSON
  ├─ generira create-pass .scr s TOČNO uključenim koracima (§3)
  ├─ create pass: accoreconsole → NETLOAD → komande → QSAVE
  ├─ verify pass: accoreconsole → NETLOAD → VERIFYBATCH (read-only, bez QSAVE)
  ├─ usporedbe (§6), pa tek onda zamjena originala uz .bak
  └─ export korak: Python sam piše imena u Excel (bez AutoCAD-a)  [xlsx_lite.py]

LayoutCrator/LayoutCreatorCore (C#, SAMO core projekt — UI plugin se NE dira)
  ├─ SETFIELDSBATCH   (port SetFieldsValue.lsp)
  ├─ SETTITLESBATCH   (port SetLayoutTitles.lsp)
  ├─ SORTTABSBATCH    (jezgra TabSort + ele:layout< komparator)
  ├─ VERIFYBATCH      (read-only dump stanja za verifikaciju)
  └─ BatchDataFile.cs (čitanje LAYOUT_BATCH_JSON preko System.Text.Json — u .NET 8 je in-box)
```

`LAYOUT_BATCH_JSON` (UTF-8, piše ga orkestrator u temp dir uz kopiju DWG-a):
```json
{
  "polja":   {"INVESTITOR": "HEP ODS", "GRADILISTE": "TS 10/0.4"},
  "naslovi": [{"layout": "1-1x1", "naslov": "Situacija", "mjerilo": "1:500"}]
}
```
(Ključevi prisutni samo za uključene korake; komande koje ne nađu svoj ključ
ispišu `ERR|` i ne diraju ništa.)

## 3. Redoslijed operacija po datoteci (create pass .scr)

Orkestrator SLAŽE .scr (ASCII-only!) samo od uključenih koraka, ovim redom:

```
(command "._NETLOAD" "<dll>")
CREATELAYOUTBATCH          ; ako koraci.layouti
SETFIELDSBATCH             ; ako koraci.polja
SETTITLESBATCH             ; ako koraci.naslovi
SORTTABSBATCH              ; ako koraci.sortiranje
_.UPDATEFIELD              ; ako koraci.polja — NAKON svih izmjena
_All
                           ; (prazna linija = kraj selekcije)
._QSAVE
```

- `UPDATEFIELD _All` je vjerni ekvivalent LISP-a (`SetFieldsValue.lsp:176`).
  Ako se na korisnikovom stroju pokaže da UPDATEFIELD u core konzoli ne postoji,
  fallback je prihvatljiv: polja se ionako reevaluiraju pri otvaranju/plotu
  (REGEN) — orkestrator tada samo logira WARN, ne ERR. **Označiti kao stavku
  ručne provjere.**
- Export se NE izvršava u AutoCAD-u: orkestrator ga radi sam (§4.4) NAKON
  uspješnog verifyja, iz verificirane liste layouta.
- Template `batch/templates/create_layouts.scr.tpl` dobiva placeholder
  `__COMMANDS__` koji orkestrator puni; verify template prelazi na VERIFYBATCH.

## 4. Specifikacija koraka — vjerni portovi LISP ponašanja

### 4.1. SETFIELDSBATCH (izvor: `AutoLisp/SetFieldsValue.lsp`)
- Ulaz: `polja` mapa iz JSON-a (orkestrator ju je pročitao iz sheeta
  `sheet_polja`, default **"Podaci"**: stupac A=ključ, B=vrijednost; redci s
  praznim A se preskaču — isto kao LISP "prazni redci"; bez limita od 200
  redaka — čita se cijeli sheet, svjesno poboljšanje).
- Ponašanje: postojeći custom ključ → prepiši; novi → dodaj. Vrijednosti su
  uvijek stringovi (Python normalizira brojeve: cjelobrojni float → "42",
  ostalo minimalni zapis s točkom).
- C# mehanika: `var b = new DatabaseSummaryInfoBuilder(db.SummaryInfo);`
  → `b.CustomPropertyTable[key] = value;` → `db.SummaryInfo = b.ToDatabaseSummaryInfo();`
  (ključevi case-sensitivno kao u LISP-u; usporedba postojanja preko
  CustomPropertyTable.Contains).
- Markeri: po retku ništa (prebučno); sažetak
  `RESULT|fields|added=N overwritten=M`.

### 4.2. SETTITLESBATCH (izvor: `AutoLisp/SetLayoutTitles.lsp`)
- Ulaz: `naslovi` lista iz JSON-a (sheet `sheet_nacrti`, default **"Nacrti"**,
  stupci A=ime layouta, B=naslov, C=mjerilo; prazan A → preskoči; A prisutan a
  B i C oba prazna → preskoči + WARN "nema podataka"; jedan od B/C prazan →
  upiši prazan string u taj atribut).
- Za svaki red: nađi layout po imenu (case-insensitive usporedba — AutoCAD
  imena layouta tako tretira); u njegovom paper space BTR-u nađi PRVI `INSERT`
  bloka **`sastAu`** (obični blok, nije dinamički; ime preko
  `DynamicBlockTableRecord`-a radi robusnosti); u njemu postavi atribute s
  tagovima **`layout_title`** i **`layout_mjerilo`** (tag usporedba
  case-insensitive, kao `vl-setattributevalue`).
- C# gotcha: `AttributeReference.TextString` postavljati s
  `HostApplicationServices.WorkingDatabase = db` postavljenim (AdjustAlignment
  inače zna koristiti krivu bazu); att otvoriti ForWrite.
- Ishodi (imenovanje iz LISP sažetka): ok / noblock (nema bloka ILI nema
  atributa) / nolayout. Markeri: `TITLES|ok|<layout>`, `TITLES|noblock|<layout>`,
  `TITLES|nolayout|<layout>`; sažetak `RESULT|titles|ok=N noblock=M nolayout=K`.

### 4.3. SORTTABSBATCH (komparator iz `ExportLayoutsToExcel.lsp`, jezgra TabSorta)
- Pattern valjanog imena (port `ele:valid-layout-p`, ekvivalent regexa
  `^\d+-\d+x\d+$`): znamenke, JEDNA crtica, znamenke, JEDNO 'x', znamenke.
- Parser (port `ele:nums`): "10-2x3" → (10, 2, 3). Komparator (port
  `ele:layout<`): sortiraj po PRVOM pa DRUGOM broju (treći se ignorira — vjerno
  LISP-u); sort mora biti STABILAN (Python/C# stabilni sortovi to daju).
- Dodjela: Model ostaje TabOrder 0; valjani layouti sortirani dobivaju 1..N;
  nevaljani (npr. CUSTOM template layouti) idu IZA njih, N+1.., u zatečenom
  međusobnom redoslijedu (stabilno). `Layout.TabOrder` ForWrite.
- Markeri: `TABORDER|<pozicija>|<ime>` za svaki paper layout u konačnom
  redoslijedu; sažetak `RESULT|sorttabs|sorted=N other=M`.
- C# parser/komparator u zasebnoj datoteci `LayoutNameOrder.cs` (koristi ju i
  VERIFYBATCH za očekivani poredak) — logika mora biti IDENTIČNA Python
  modulu `batch/layout_names.py` (§4.5); oba nose komentar da su port
  `ele:nums`/`ele:layout<` i da se mijenjaju samo zajedno.

### 4.4. Export u Excel (izvor: `AutoLisp/ExportLayoutsToExcel.lsp`; IZVODI GA PYTHON)
- Nakon USPJEŠNOG verifyja: uzmi verificiranu listu layouta (iz VERIFYBATCH
  ispisa), filtriraj pattern `^\d+-\d+x\d+$`, sortiraj komparatorom iz
  `layout_names.py`, upiši imena u **stupac A, redci 1..N** sheeta
  `sheet_nacrti` (default "Nacrti").
- Vjerno LISP-u: NE briši sadržaj ispod retka N (LISP PutCell samo prepisuje
  A1..AN). Ako je u stupcu A ispod N bilo ne-praznih ćelija → WARN u log/CSV.
- Ako Excel datoteka ne postoji → kreiraj minimalni novi .xlsx s tim sheetom.
  Ako postoji a sheet ne postoji → dodaj novi sheet (workbook.xml + rels +
  [Content_Types].xml + novi sheet part). Ako je datoteka zaključana (otvorena
  u Excelu) → piši u temp pa `os.replace` s 3 pokušaja; ako ne uspije → ERR
  redak u CSV (DWG izmjene ostaju valjane — export je jedini Excel-write korak).

### 4.5. Python moduli (novi, SAMO stdlib, s unit testovima)
- `batch/layout_names.py` — `nums(str)`, `valid_layout(str)`, `layout_sort_key`
  (port ele funkcija). Testovi s primjerima iz LISP komentara ("3-1x1", "10-2x3",
  odbacivanje "1-2-3", "A4L", "1x1", "3-1x1x2"...).
- `batch/xlsx_lite.py` — `read_sheet(path, sheet) -> list[list[str]]`
  (workbook.xml → sheet mapping, sharedStrings, inlineStr, brojevi
  normalizirani kao u §4.1) i `write_column_a(path, sheet, values) -> upozorenja`
  (§4.4 semantika). Testovi: modul sam generira minimalni xlsx pa ga čita
  natrag (round-trip), plus čitanje xlsx-a sa sharedStrings i s inlineStr.

### 4.6. VERIFYBATCH (novo, read-only)
Ispisuje SVE potrebno za usporedbe, bez ikakvih izmjena i bez QSAVE:
- `VLAYOUT|<taborder>|<ime>` za svaki paper layout,
- `VPROP|<ključ>=<vrijednost>` za svaki custom drawing property,
- `VTITLE|<layout>|<title>|<mjerilo>` za svaki layout koji ima `sastAu`
  (vrijednosti atributa layout_title/layout_mjerilo).
Vrijednosti mogu sadržavati `|` → orkestrator parsira s ograničenim brojem
splitova (`split("|", maxsplit)`), s desna kad treba.

## 5. Proširenje `batch/config.json` (unatrag kompatibilno!)

```json
{
  "...postojeća polja ostaju ista...": "",
  "excel": "C:\\Projekti\\zajednicki_podaci.xlsx",
  "sheet_polja": "Podaci",
  "sheet_nacrti": "Nacrti",
  "koraci": { "layouti": true, "polja": false, "naslovi": false,
              "sortiranje": false, "export": false },
  "projekti": [
    "C:\\Projekti\\P100\\elektro.dwg",
    { "dwg": "C:\\Projekti\\P101\\elektro.dwg",
      "excel": "C:\\Projekti\\P101\\podaci.xlsx" }
  ]
}
```
- `koraci` izostavljen → `{"layouti": true}`, sve ostalo false (stari config
  radi identično kao prije).
- `excel` na razini projekta (objekt-zapis) nadjačava globalni; koraci
  polja/naslovi/export bez dostupnog Excela → preflight ERR za te datoteke
  (jasna poruka), layouti/sortiranje se svejedno smiju izvršiti.
- `resolve_projects` prihvaća string ILI objekt (backward compat).

## 6. Verifikacija po koraku i CSV

Nakon create passa orkestrator scrape-a postojeće (`CREATED|`, `RESULT|created=`)
i nove markere. Verify pass (VERIFYBATCH na KOPIJI, prije zamjene originala)
mora potvrditi:
- **layouti**: svako `CREATED|` ime postoji u `VLAYOUT|` (postojeća logika,
  samo izvor postaje VERIFYBATCH umjesto LISP dictsearch),
- **polja**: svaki poslani ključ postoji u `VPROP|` s točnom vrijednošću,
- **naslovi**: za svaki `TITLES|ok|` layout `VTITLE|` vrijednosti == poslane,
- **sortiranje**: `VLAYOUT|` poredak == očekivani (isti komparator u Pythonu).
Bilo koji mismatch → ERR, original NETAKNUT, stdout dump u log.

CSV `rezultati_*.csv` dobiva stupac **`korak`**
(layouti|polja|naslovi|sortiranje|export|verify) između `datoteka` i `layout`;
`layout` stupac za ne-layout retke nosi ključ/ime (npr. ime propertyja kod
greške). Delimiter `;` i UTF-8 BOM ostaju.

## 7. GUI (`batch/batch_gui.py`) — minimalno proširenje
Checkboxovi za 5 koraka + polja za `excel`, `sheet_polja`, `sheet_nacrti`
(s Browse za excel). Config round-trip mora ostati kompatibilan s CLI-jem.
Ne mijenjati threading model ni tok obrade — GUI samo puni prošireni config.

## 8. AutoLisp repo (ista grana)
- SAMO dokumentacija: u `README.md` (sekcija layout alata) dodati napomenu da
  `SetFieldsValue`, `SetLayoutTitle`, `ele` i sortiranje tabova imaju headless
  batch portove u DwgLayoutCreator repou (interaktivne LISP verzije ostaju
  referentne i dalje rade). NIKAKVE izmjene .lsp datoteka → nema
  ACADDOC/ACADLT sinkronizacije. Ne dirati Lee Mac datoteke.

## 9. Konvencije koje MORAŠ poštovati
- Hrvatski: svi komentari, poruke, README tekstovi; machine-markeri ostaju
  engleski (`RESULT|fields|...`) jer su API orkestratora.
- NE dirati: interaktivni plugin `LayoutCrator/LayoutCreator` (osim ako treba
  `#if CORECONSOLE` u DIJELJENIM datotekama — tada UI put mora ostati
  bajt-ekvivalentan ponašanju), zamrznuti `DwgLayoutCreator/` (ACadSharp),
  `poc/`.
- Python: SAMO stdlib. C#: bez NuGet paketa (System.Text.Json je in-box u .NET 8).
- Nove .cs datoteke idu u `LayoutCrator/LayoutCreatorCore/` (nisu dijeljene s
  UI pluginom — kompajliraju se samo u core build, kao CreateLayoutBatchCommand).
- .scr sadržaj ASCII-only.
- Commit poruke: kratke, hrvatski, `feat:`/`fix:`/`refactor:`/`docs:` prefiks;
  logične cjeline (moduli+testovi, C# komande, orkestrator, GUI, docs). Bez PR-a.

## 10. Datoteke — popis i redoslijed implementacije

1. `batch/layout_names.py` + `batch/tests/test_layout_names.py` (pokreni testove!)
2. `batch/xlsx_lite.py` + `batch/tests/test_xlsx_lite.py` (pokreni testove!)
3. `LayoutCrator/LayoutCreatorCore/LayoutNameOrder.cs`
4. `LayoutCrator/LayoutCreatorCore/BatchDataFile.cs`
5. `LayoutCrator/LayoutCreatorCore/SetFieldsBatchCommand.cs`
6. `LayoutCrator/LayoutCreatorCore/SetTitlesBatchCommand.cs`
7. `LayoutCrator/LayoutCreatorCore/SortTabsBatchCommand.cs`
8. `LayoutCrator/LayoutCreatorCore/VerifyBatchCommand.cs`
9. `batch/templates/create_layouts.scr.tpl` (placeholder `__COMMANDS__`),
   `batch/templates/verify_layouts.scr.tpl` (→ NETLOAD + VERIFYBATCH)
10. `batch/batch_layouts.py` (config shema, JSON data file, .scr slaganje,
    novi markeri+verify usporedbe, export korak, CSV stupac `korak`)
11. `batch/batch_gui.py` (checkboxovi + excel polja)
12. `batch/config.json` (prošireni primjer), `README.md` (nove sekcije:
    koraci, config polja, troubleshooting za UPDATEFIELD/xls/locked-Excel)
13. AutoLisp `README.md` napomena
14. Sve committati i pushati na `claude/fable-layout-automation-rv8gu2`
    (oba repoa; `git push -u origin <grana>`, kod mrežnih grešaka retry 2s/4s/8s/16s)

## 11. Verifikacija — što je moguće ovdje, a što kod korisnika

**Ovdje (obavezno napraviti i izvijestiti):**
- `python -m py_compile` na svim .py; `python -m unittest` za `batch/tests/*`
  (layout_names, xlsx_lite round-trip) — ovi testovi RADE bez Windowsa;
- ručna provjera generiranog .scr sadržaja (dry-run ispis za sve kombinacije koraka);
- config round-trip (stari config bez novih polja mora proći identično kao prije);
- C#: NEMA build okruženja (nema dotnet-a ni AutoCAD DLL-ova u sandboxu) —
  navedi to izričito; kod piši konzervativno po uzoru na postojeće komande.

**Kod korisnika (napiši u README "Ručna provjera" korake):**
1. `dotnet build LayoutCrator/LayoutCreatorCore -c Release` (mora proći bez grešaka).
2. Batch s `koraci: {layouti:true, sortiranje:true}` na kopiji test projekta →
   tabovi numerički poredani, CSV bez ERR.
3. Excel s "Podaci" (2-3 para) i "Nacrti" (naslov+mjerilo za 2 layouta) →
   batch s polja+naslovi → u AutoCAD-u provjeri custom properties (DWGPROPS),
   naslove u sastavnici i da FIELD-ovi pokazuju nove vrijednosti (UPDATEFIELD
   stavka: ako WARN u logu, javi — ugrađen je REGEN fallback).
4. Export: batch s export:true → stupac A sheeta "Nacrti" == sortirana imena.
5. Idempotentnost: isti batch 2× → drugi prolaz 0 promjena, 0 ERR.
