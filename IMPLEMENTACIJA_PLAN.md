# IMPLEMENTACIJA_PLAN.md — Batch generiranje layouta preko više DWG projekata

> Plan za Opusa. Nastao nakon dijagnoze sva tri repoa (AutoLisp, DwgLayoutCreator,
> EE-Python-Tools). Cilj: headless batch kreiranje paperspace layouta (viewport +
> sastavnica + page setup) za više DWG projekata odjednom, bez ručnog otvaranja
> svakog crteža.

---

## 1. Dijagnoza — zašto su prethodni pokušaji pukli

### 1.1. Što postoji (stanje repoa)

| Artefakt | Repo | Status |
|---|---|---|
| `c:createlayout` (createlayoutV2.lsp) | AutoLisp | **Radi.** Interaktivni alat u AutoCAD-u: skenira model space za blokove A4L…A0P + CUSTOM1-8, čita broj layouta iz atributa (BR_N ili zadnji atribut), kreira layout, page setup preko `-pagesetup`, mview, umeće `sastAu.dwg` sastavnicu i puni BR_N/BR_L/BR_LI. Referentna implementacija ponašanja. |
| `LayoutCrator/LayoutCreator` (.NET plugin, NETLOAD) | DwgLayoutCreator | **Najzrelija implementacija.** Komanda `CREATELAYOUT` — neinteraktivna, page setup preko `PlotSettingsValidator` s 4-razinskim matchanjem canonical media imena, dinamički blokovi preko `DynamicBlockTableRecord`, pravi `GeometricExtents`, dijagnostičke komande `DIAGLAYOUTS` i `LISTMEDIA`. Cilja AutoCAD 2025 (`AcadDir` u csproj). |
| `DwgLayoutCreator` (standalone, ACadSharp 3.5.7) | DwgLayoutCreator | **Slijepa ulica** — vidi 1.2. |
| `syncProperties.py`, `mcp_autocad/bridge.py` | EE-Python-Tools | Rade. Dokazana COM infrastruktura (`win32com`, `_com_retry` za `RPC_E_CALL_REJECTED`, GetActiveObject/Dispatch fallback). |
| `dwg_props_helper` (ACadSharp) | EE-Python-Tools | Radi, ali **samo čita** DWG — zato i radi. |

### 1.2. Točan uzrok neuspjeha

Pokušaji "uredi DWG kao datoteku, bez AutoCAD-a" (ezdxf u Pythonu, pa ACadSharp u
C#-u) padaju **svi na istoj klasi problema**: kreiranje layouta u ovom workflowu
nije uređivanje datoteke, nego ovisi o živim AutoCAD servisima. Konkretni dokazi
u kodu standalone alata:

1. **Pisanje DWG-a je minsko polje.** `DwgLayoutCreator/TitleBlockInserter.cs:79-95`
   dokumentira workaround za NRE bug u ACadSharp 3.5.7 `DwgObjectWriter.writeCommonAttData`
   (MultiLine atributi → morali su se degradirati u SingleLine da writer ne padne).
   ezdxf uopće ne piše DWG — treba ODA File Converter round-trip
   (DWG→DXF→izmjena→DXF→DWG), što je sporo, gubi podatke i dodaje vanjsku ovisnost.

2. **Dinamički blokovi (CUSTOM1-8) se ne vide.** Standalone scanner
   (`DwgLayoutCreator/ModelSpaceScanner.cs:32`) matcha `ins.Block?.Name` — ali
   instanca dinamičkog bloka pokazuje na anonimni `*U##` block record, pa se
   CUSTOM blokovi tiho preskaču. Plugin verzija je isti problem morala riješiti
   preko `DynamicBlockTableRecord` (AutoCAD API — nema ga ni ezdxf ni ACadSharp).

3. **Nema PlotSettingsValidatora.** Standalone hardkodira plotter, margine i
   canonical media imena (`LayoutFactory.cs:11-21`, komentar: "nemamo PSV pa
   hardcode margine — AutoCAD će prepoznati svoj PMP pri otvaranju" — nada, ne
   garancija). Page setup ispravno radi samo validacijom protiv stvarnog
   `DWG To PDF.pc3` + korisnikovog PMP-a (25 mm uvez!), što izvan AutoCAD-a ne postoji.
   Čak i s pravim API-jem je trebalo 4-razinsko matchanje media imena
   (`PageSetupConfigurator.FindCanonicalByLocaleName`) — ezdxf tu piše stringove naslijepo.

4. **Nema GeometricExtents.** Standalone aproksimira bounding box iz insert
   pointova djece s fallbackom ±50 mm (`ModelSpaceScanner.cs:69-110`) → krivi zoom
   viewporta. Plugin koristi pravi `bref.GeometricExtents`.

**Zaključak:** krivac nije Python ni ezdxf kao takav, nego arhitektura. Sve što
dira layout/page-setup/dinamičke blokove mora izvršavati **pravi AutoCAD engine**.
Batch dimenziju treba riješiti orkestracijom oko njega, ne zaobilaženjem.

### 1.3. Preduvjet koji nedostaje u repou

U repoima **nema nijednog primjera DWG projekta ni sastavnice**. Prije početka
implementacije korisnik mora u lokalni folder (npr. `C:\BatchLayoutTest\`) staviti:
2-3 reprezentativna projekta (s A4L/A3L… i barem jednim CUSTOM blokom + template
layoutom) i `sastAu.dwg`. Bez toga se koraci ne mogu verificirati.

---

## 2. Evaluacija alata (za ovaj slučaj: Windows, AutoCAD 2025, DWG, dinamički blokovi, PC3/PMP page setup)

| Pristup | Ocjena | Ključna ograničenja za NAŠ slučaj |
|---|---|---|
| **ezdxf (+ ODA konverzija)** | ❌ odbaciti | Ne piše DWG (samo DXF → ODA round-trip); nema effective name dinamičkih blokova → CUSTOM1-8 nevidljivi; nema PSV → page setup naslijepo; bbox aproksimacija. Identične rupe zbog kojih je i ACadSharp pokušaj stao. |
| **ACadSharp standalone (postojeći kod)** | ❌ zamrznuti | Vidi 1.2. Ostaviti samo za **read-only** (kao `dwg_props_helper`). |
| **COM automatizacija (win32com/pyautocad)** | ⚠️ fallback | Native i sve radi (dokazano u `syncProperties.py`), ali: puni AutoCAD po dokumentu (sporo), `RPC_E_CALL_REJECTED` flakiness (već trebaju retryji), modalni dijalozi znaju blokirati batch, nema paralelizma. Dobro kao plan B i za vizualnu verifikaciju. |
| **accoreconsole.exe + postojeći .NET plugin** | ✅ **ODABRANO** | Pravi AutoCAD core headless (dolazi uz AutoCAD 2025, ne uz LT). Izvršava već izdebugirani `CREATELAYOUT` (PSV, dinamički blokovi, extents). Brz (nema UI), paralelizabilan, exit codeovi. Caveati (rješivi, vidi korak 1): plugin treba prebaciti s `acmgd` na core-safe API; ActiveX/vla-* NE radi u core konzoli (zato ne može LISP verzija); NETLOAD traži TRUSTEDPATHS. |
| **Sheet Set Manager** | ❌ nije fit | Organizira POSTOJEĆE layoute u sheet setove; ne zna skenirati model-space blokove, kreirati viewporte ni puniti sastavnicu. Rješava krivi problem. |

**Odabir: Python orkestrator → accoreconsole.exe → NETLOAD core-verzije postojećeg
.NET plugina → `CREATELAYOUT`.** Jedini pristup koji ponovno koristi već
izdebugiranu logiku (page setup, dinamički blokovi, sastavnica) i daje headless
batch bez prepisivanja domenske logike. COM ostaje dokumentirani fallback.

---

## 3. Arhitektura

```
┌─────────────────────────────────────────────────────────────────┐
│ batch/batch_layouts.py  (Python 3.10+, SAMO stdlib)             │
│  1. učita batch/config.json                                     │
│  2. za svaki DWG: kopija → temp                                 │
│  3. generira .scr iz templatea (ASCII!)                         │
│  4. subprocess: accoreconsole /i kopija.dwg /s skripta.scr      │
│  5. parsa stdout (INFO|/OK|/WARN|/ERR| konvencija već postoji)  │
│  6. verifikacijski pass (layoutlist → txt)                      │
│  7. uspjeh → zamijeni original (uz .bak); neuspjeh → skip + log │
│  8. summary: logs/run_<timestamp>.log + rezultati.csv           │
└───────────────┬─────────────────────────────────────────────────┘
                │ subprocess, timeout, exit code
┌───────────────▼─────────────────────────────────────────────────┐
│ accoreconsole.exe (AutoCAD 2025 core, headless)                 │
│  .scr: SECUREREMOTEACCESS/putanje → NETLOAD LayoutCreatorCore   │
│        → CREATELAYOUTBATCH → QSAVE                              │
└───────────────┬─────────────────────────────────────────────────┘
                │ in-process .NET API
┌───────────────▼─────────────────────────────────────────────────┐
│ LayoutCreatorCore.dll (refactor postojećeg LayoutCreator)       │
│  ModelSpaceScanner → LayoutBuilder → PageSetupConfigurator      │
│  → TitleBlockInserter  (logika NEPROMIJENJENA)                  │
└─────────────────────────────────────────────────────────────────┘
```

**Tok podataka o layoutima:** parametri layouta (format papira, broj, položaj)
**ostaju u DWG-u** kao i dosad — blokovi `A4L…A0P`/`CUSTOM1-8` s atributom broja
(BR_N ili zadnji atribut). To je postojeća konvencija koju koriste i LISP i
plugin; config NE duplicira te podatke, samo orkestrira (koje datoteke, putanje
alata, sastavnica, timeout).

**Logging:** plugin već ispisuje statusne poruke; standardizirati na
`OK|`/`WARN|`/`ERR|`/`INFO|` prefikse (konvencija iz standalone `Program.cs`).
Orkestrator ih parsa iz stdout-a i agregira u CSV: jedna linija po
(datoteka, layout, status, poruka).

---

## 4. Koraci implementacije (redoslijed obavezan)

### Korak 0 — PoC na korisnikovom stroju (VEĆ NAPISAN: `poc/poc_accoreconsole.py`)
Korisnik pokreće `python poc\poc_accoreconsole.py "C:\BatchLayoutTest\projekt1.dwg"`.
- **DoD:** ispis `PASS` na 2-3 reprezentativna projekta. Ako FAIL → stani,
  dijagnosticiraj s korisnikom prije bilo kakvog daljnjeg koda (ne nastavljati naslijepo).

### Korak 1 — `LayoutCreatorCore`: core-console-kompatibilan build plugina
Novi projekt `LayoutCrator/LayoutCreatorCore/LayoutCreatorCore.csproj` (net8.0-windows,
x64), datoteke **dijeli** s postojećim pluginom (link na iste .cs), razlike:
- Referencirati samo `accoremgd.dll` + `acdbmgd.dll` iz `$(AcadDir)` (bez `acmgd`).
- `Application` → `Autodesk.AutoCAD.ApplicationServices.Core.Application`
  (conditional compile simbol `CORECONSOLE` ili alias-using; NE forkati logiku).
- `Application.SetSystemVariable("PSLTSCALE"/"LTSCALE"/"MSLTSCALE", …)` zamijeniti
  core-safe ekvivalentima na `Database` (npr. `db.Psltscale`, `db.Ltscale`,
  `db.Msltscale`) — provjeriti točna imena propertyja pri implementaciji.
- `doc.SendStringToExecute("regenall…")` maknuti/zamijeniti (u headless kontekstu
  regen nije bitan; AutoCAD regenerira pri otvaranju).
- Nova komanda `CREATELAYOUTBATCH`: ista logika kao `CREATELAYOUT`, ali putanju
  sastavnice čita iz env varijable `LAYOUT_SAST_PATH` (postavlja je orkestrator),
  s fallbackom na `HostApplicationServices.FindFile("sastAu.dwg")` — jer support
  file search path u core konzoli NIJE isti kao u punom AutoCAD-u.
- **DoD:** `dotnet build -c Release` prolazi bez upozorenja o missing referencama;
  na korisnikovom stroju `accoreconsole /i test.dwg /s netload_diag.scr` uspješno
  NETLOAD-a DLL i `DIAGLAYOUTS` ispiše blokove (uključivo CUSTOM dinamičke).

### Korak 2 — Orkestrator `batch/batch_layouts.py`
Python, **samo stdlib** (argparse, json, subprocess, shutil, pathlib, csv, logging,
concurrent.futures). Ponašanje:
- `--config batch/config.json`; `--dry-run` (samo ispiši što bi se radilo);
  `--jobs N` (paralelno, default 1; svaka instanca accoreconsole je neovisna).
- Po datoteci: preskoči ako postoji `.dwl`/`.dwl2` (otvorena u AutoCAD-u) ili je
  read-only → `ERR` u CSV. Kopija u temp, `.scr` generiran iz
  `batch/templates/create_layouts.scr.tpl` (ASCII-only — bez č/ć/š/đ/ž u .scr!).
- `subprocess.run(..., timeout=cfg)` + kill pri timeoutu; exit code != 0 → ERR.
- Nakon create passa: verifikacijski pass (kao u PoC-u — `(layoutlist)` u txt) i
  usporedba s očekivanim brojem kandidata (parsano iz stdout-a create passa).
- Uspjeh → original preimenuj u `<ime>.dwg.bak`, kopiju vrati na mjesto originala;
  neuspjeh → original NETAKNUT.
- **DoD:** na 2-3 test projekta kreira layoute; ponovljeno pokretanje kreira 0
  novih (idempotentnost — postojeći layouti se preskaču, semantika već u pluginu);
  `rezultati.csv` točno odražava stdout; original ostaje netaknut kad se ubije
  proces usred rada.

### Korak 3 — NETLOAD security + SCR template
- `.scr` template redoslijed: `NETLOAD` s punom putanjom DLL-a → `CREATELAYOUTBATCH`
  → `._QSAVE`.
- Ako `SECURELOAD` blokira NETLOAD: dokumentirati (i orkestrator neka detektira
  poruku u stdout-u) da se DLL folder doda u `TRUSTEDPATHS` **ili** pokretati
  accoreconsole s `/product ACAD /language en-US` + profilom gdje je folder trusted.
  Rješenje utvrditi na korisnikovom stroju — ne pogađati.
- **DoD:** batch prolazi na svježem accoreconsole procesu bez ikakve interakcije.

### Korak 4 — Rubni slučajevi i hardening (vidi §6)
- **DoD:** svaki slučaj iz tablice u §6 ima implementirano ponašanje i ručno
  izazvan test (npr. zaključaj DWG pa pokreni batch).

### Korak 5 — Integracija i dokumentacija
- `launchers/Batch Layouti.bat` u **EE-Python-Tools** repou (isti branch
  `claude/cad-layout-automation-plan-l4dead`), po uzoru na postojeće launchere.
- README sekcija u DwgLayoutCreator (korištenje, config polja, troubleshooting).
- **DoD:** launcher radi s network share deploy konvencijom (`deploy.ps1` flow).

**Izvan scopea (svjesno):** plot u PDF iz batcha — `PrintOrganizer.py` i postojeći
publish workflow to već pokrivaju; dodati tek kad layout batch bude stabilan.

---

## 5. Struktura direktorija i config

```
DwgLayoutCreator/
├── LayoutCrator/
│   ├── LayoutCreator/            # postojeći UI plugin (NE dirati ponašanje)
│   └── LayoutCreatorCore/        # NOVO — core console build (Korak 1)
├── batch/                        # NOVO (Korak 2)
│   ├── batch_layouts.py
│   ├── config.json               # primjer, committati
│   ├── templates/
│   │   ├── create_layouts.scr.tpl
│   │   └── verify_layouts.scr.tpl
│   └── logs/                     # gitignore
├── poc/
│   └── poc_accoreconsole.py      # već postoji (Korak 0)
└── IMPLEMENTACIJA_PLAN.md
```

`batch/config.json`:
```json
{
  "accoreconsole": "C:\\Program Files\\Autodesk\\AutoCAD 2025\\accoreconsole.exe",
  "plugin_dll":    "C:\\...\\LayoutCreatorCore\\bin\\Release\\LayoutCreatorCore.dll",
  "sastavnica":    "\\\\server\\predlosci\\sastAu.dwg",
  "timeout_s":     300,
  "backup":        true,
  "jobs":          1,
  "projekti": [
    "C:\\Projekti\\P100\\elektro.dwg",
    "C:\\Projekti\\P101\\elektro.dwg"
  ]
}
```
Polja `projekti` alternativno: `{"folder": "C:\\Projekti", "pattern": "**/elektro*.dwg"}`
(implementirati oba). Sve ostalo o layoutima čita se iz DWG blokova (§3).

**Točne verzije:** .NET SDK 8.0.x; AutoCAD 2025 (accoreconsole iz instalacije;
AutoCAD LT NEMA accoreconsole ni .NET — batch se vrti samo na stroju s punim
AutoCAD-om); Python ≥ 3.10, bez third-party paketa; AutoCAD DLL reference iz
install foldera (kao u postojećem csproj-u, `AcadDir` property override za
druge verzije). ACadSharp 3.5.7 ostaje SAMO u read-only `dwg_props_helper`.

---

## 6. Rubni slučajevi

| Slučaj | Ponašanje |
|---|---|
| DWG otvoren u AutoCAD-u (`.dwl` postoji) | Preskoči, `ERR|zaključan` u CSV, nastavi s ostalima. |
| Read-only / nedostupna putanja | Preskoči + ERR. |
| Layout ime `X` (nenumerirano) | Preskoči taj blok (postojeća semantika u pluginu). |
| Layout već postoji | Preskoči (idempotentnost; postojeća semantika). |
| CUSTOM blok bez template layouta istog imena | WARN + preskoči (postojeća semantika). |
| `sastAu.dwg` nedostupan | Fatalno za tu datoteku PRIJE ikakve izmjene — ERR, original netaknut. |
| accoreconsole timeout/hang | `subprocess` timeout → kill → ERR; original netaknut (radili smo na kopiji). |
| Crash usred pisanja | Original se mijenja tek NAKON uspješnog verify passa; uvijek `.bak`. |
| Stariji DWG format | accoreconsole otvara sve; spremati u zatečenom formatu (QSAVE), ne raditi SAVEAS konverzije. |
| Ne-ASCII znakovi | U `.scr` i imenima layouta iz atributa — imena layouta dolaze iz DWG-a pa mogu imati č/ć; `.scr` template mora ostati ASCII, a imena prolaze kroz plugin (in-process .NET), ne kroz .scr → OK. Log/CSV pisati kao UTF-8 s BOM (Excel). |
| Educational stamp / recover-prompt crteži | Detektirati po stdout poruci → ERR + uputa da se sredi ručno. |
| Duplikatni brojevi layouta u istom crtežu | Prvi pobjeđuje, drugi se preskače kao "već postoji" + WARN (dokumentirati). |

---

## 7. Test plan

1. **PoC (Korak 0):** PASS na 2-3 korisnikova projekta — dokazuje headless chain.
2. **DIAGLAYOUTS headless (Korak 1):** ispis mora pokazati iste kandidate koje
   vidi interaktivni `DIAGLAYOUTS` u punom AutoCAD-u na istom crtežu (ručna
   usporedba jednom), posebno CUSTOM dinamičke blokove.
3. **E2E (Korak 2):** batch na kopijama 2-3 projekta → za svaki: broj kreiranih
   layouta == broj numeriranih blokova u model spaceu; verify pass potvrđuje imena.
4. **Idempotentnost:** isti batch 2× zaredom → drugi prolaz 0 kreiranih, 0 grešaka.
5. **Vizualna provjera (korisnik, jednom po formatu):** otvoriti jedan generirani
   layout u AutoCAD-u — papir/orijentacija točni, margine (25 mm uvez) poštovane,
   viewport zumira na blok, sastavnica na desnom rubu s ispravnim BR_N/BR_L/BR_LI,
   `monochrome.ctb` aktivan, plot preview izgleda kao kod LISP verzije.
6. **Negativni testovi:** zaključan DWG, krivi path sastavnice, ubijen proces —
   original uvijek netaknut.
7. Napomena: u repou nema DWG fixtura pa su testovi 2-6 izvodivi samo na
   korisnikovom stroju — svaki korak koji to zahtijeva jasno označiti korisniku
   i tražiti potvrdu rezultata, ne tvrditi "testirano" bez pokretanja.

---

## 8. Čega se NE hvatati (naučene lekcije)

- **Ne** vraćati se na ezdxf/ACadSharp za PISANJE DWG-a — to je uzrok prošlih
  neuspjeha (§1.2), ne rješenje.
- **Ne** prepisivati domensku logiku ispočetka — `LayoutBuilder`/`PageSetupConfigurator`/
  `TitleBlockInserter`/`ModelSpaceScanner` iz plugina su već izdebugirani
  (eNotInDatabase redoslijed, orientation bug, canonical media matchanje, paper-boundary
  viewport #1) — refactor je samo core-console adaptacija.
- **Ne** koristiti vla-*/vlax-* LISP u accoreconsole skriptama (ActiveX tamo ne radi).
- **Ne** dirati interaktivni `c:createlayout` u AutoLisp repou — ostaje alat za
  pojedinačne crteže i za LT korisnike.

## 9. Rizici / otvorena pitanja (razriješiti u Koraku 0-1, ne kasnije)

1. Ponaša li se `PlotSettingsValidator` + `GetPlotDeviceList` u core konzoli
   identično (vidi li `DWG To PDF.pc3` i korisnikov PMP)? → provjeriti `LISTMEDIA`
   headless; `ResolvePlotter` već ima fallbackove.
2. `SECURELOAD`/`TRUSTEDPATHS` politika na korisnikovom stroju (Korak 3).
3. Support path za `sastAu.dwg` u core konzoli → zato env varijabla u Koraku 1.

---

## 10. Uputa za pokretanje implementacije (zalijepi Opusu)

> U repou DwgLayoutCreator na branchu `claude/cad-layout-automation-plan-l4dead`
> nalazi se `IMPLEMENTACIJA_PLAN.md` — implementiraj ga korak po korak, redoslijedom
> iz §4, bez preskakanja "definition of done" provjera. Prvo me provedi kroz Korak 0
> (PoC `poc/poc_accoreconsole.py` na mojim primjerima u `C:\BatchLayoutTest\`) i stani
> ako ne dobijem PASS. Ne uvodi nove biblioteke mimo plana (nikakav ezdxf/ACadSharp
> za pisanje DWG-a), ponovno iskoristi postojeći kod iz `LayoutCrator/LayoutCreator`,
> a sve što zahtijeva moj AutoCAD izričito označi kao korak koji ja pokrećem i čekaj
> moju potvrdu rezultata prije nastavka.
