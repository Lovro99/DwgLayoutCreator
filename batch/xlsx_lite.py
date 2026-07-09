#!/usr/bin/env python3
"""
xlsx_lite.py — citanje i pisanje .xlsx/.xlsm datoteka BEZ Excela i BEZ vanjskih
paketa (SAMO stdlib: zipfile + xml.etree). Zamjena za ActiveX GetExcel.lsp koji
u accoreconsole ne postoji (plan §1, opcija b).

Javne funkcije:
  - read_sheet(path, sheet) -> list[list[str]]
        Procita sheet u 2D listu stringova. Brojevi normalizirani (cijeli float
        -> "42", ostalo minimalni zapis s tockom). sharedStrings i inlineStr
        oba podrzana.
  - write_column_a(path, sheet, values) -> list[str]
        Upise values u stupac A, redci 1..N zadanog sheeta (NE brise sadrzaj
        ispod retka N — vjerno LISP PutCell). Vraca listu upozorenja (npr.
        ne-prazne celije u stupcu A ispod retka N). Atomican zapis (temp +
        os.replace, 3 pokusaja). Ako datoteka/sheet ne postoji — kreira ih.

Ogranicenja (plan §1): podrzani su SAMO .xlsx i .xlsm (ZIP). Stari binarni .xls
NIJE podrzan (podigni ValueError s porukom "konvertiraj u .xlsx"). Kod .xlsm se
vbaProject.bin ne dira (svi ostali dijelovi ZIP-a kopiraju se bajt-za-bajt).
Vrijednosti se pisu kao inline stringovi (t="inlineStr") pa se sharedStrings.xml
ne dira.

SAMO stdlib.
"""

from __future__ import annotations

import os
import re
import time
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

# OOXML namespace URI-ovi.
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_CELL_REF_RE = re.compile(r"^([A-Za-z]+)([0-9]+)$")


# ---------------------------------------------------------------------------
# pomocno: reference celija <-> indeksi
# ---------------------------------------------------------------------------

def _col_to_index(letters: str) -> int:
    """"A" -> 0, "B" -> 1, "AA" -> 26 (0-based)."""
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _index_to_col(idx: int) -> str:
    """0 -> "A", 26 -> "AA" (0-based)."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _split_ref(ref: str) -> tuple[int, int]:
    """"B12" -> (col_index=1, row=12). row je 1-based kao u Excelu."""
    m = _CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f"neispravna referenca celije: {ref!r}")
    return _col_to_index(m.group(1)), int(m.group(2))


def _local(tag: str) -> str:
    """Vrati lokalno ime XML taga bez namespace prefiksa ({ns}name -> name)."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _normalize_number(text: str) -> str:
    """Cijeli float -> "42", ostalo minimalni zapis s tockom (plan §4.1)."""
    try:
        f = float(text)
    except ValueError:
        return text
    if f.is_integer():
        return str(int(f))
    return repr(f)


# ---------------------------------------------------------------------------
# citanje strukture workbooka
# ---------------------------------------------------------------------------

def _check_zip(path: str) -> None:
    if path.lower().endswith(".xls") and not path.lower().endswith(("xlsx", "xlsm")):
        raise ValueError(
            f"stari binarni .xls nije podrzan: {path} — konvertiraj u .xlsx")


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Procita xl/sharedStrings.xml (ako postoji) u listu stringova."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    result: list[str] = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        result.append(_collect_text(si))
    return result


def _collect_text(elem: ET.Element) -> str:
    """Skupi tekst svih <t> potomaka (rich-text run-ovi -> spojeni)."""
    parts = []
    for t in elem.iter():
        if _local(t.tag) == "t":
            parts.append(t.text or "")
    return "".join(parts)


def _sheet_name_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Vrati {ime_sheeta: put_do_sheet_partu} iz workbook.xml + rels."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = _read_workbook_rels(zf)
    mapping: dict[str, str] = {}
    for sheets in wb:
        if _local(sheets.tag) != "sheets":
            continue
        for sheet in sheets:
            if _local(sheet.tag) != "sheet":
                continue
            name = sheet.get("name")
            rid = sheet.get(f"{{{_NS_REL}}}id")
            if name is None or rid is None or rid not in rels:
                continue
            mapping[name] = rels[rid]
    return mapping


def _read_workbook_rels(zf: zipfile.ZipFile) -> dict[str, str]:
    """Vrati {rId: puni_put_do_partu} iz xl/_rels/workbook.xml.rels."""
    rels_name = "xl/_rels/workbook.xml.rels"
    result: dict[str, str] = {}
    if rels_name not in zf.namelist():
        return result
    root = ET.fromstring(zf.read(rels_name))
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid is None or target is None:
            continue
        result[rid] = _resolve_target(target)
    return result


def _resolve_target(target: str) -> str:
    """Target iz workbook.xml.rels je relativan na xl/ (npr. "worksheets/sheet1.xml")
    ili apsolutan ("/xl/..."). Vrati puni put unutar ZIP-a."""
    if target.startswith("/"):
        return target.lstrip("/")
    return "xl/" + target


def _find_sheet_part(zf: zipfile.ZipFile, sheet: str) -> str | None:
    """Nadi sheet part po imenu (exact pa case-insensitive fallback)."""
    mapping = _sheet_name_map(zf)
    if sheet in mapping:
        return mapping[sheet]
    lower = {k.lower(): v for k, v in mapping.items()}
    return lower.get(sheet.lower())


# ---------------------------------------------------------------------------
# read_sheet
# ---------------------------------------------------------------------------

def read_sheet(path: str, sheet: str) -> list[list[str]]:
    """Procitaj sheet u 2D listu stringova. Redci 1..maxRow, svaki popunjen do
    najveceg koristenog stupca (prazne celije = ""). Podize FileNotFoundError
    ako datoteka ne postoji, ili ValueError ako sheet ne postoji / .xls."""
    _check_zip(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Excel datoteka ne postoji: {path}")
    with zipfile.ZipFile(path) as zf:
        part = _find_sheet_part(zf, sheet)
        if part is None or part not in zf.namelist():
            raise ValueError(f"sheet '{sheet}' ne postoji u {path}")
        shared = _read_shared_strings(zf)
        root = ET.fromstring(zf.read(part))
        return _parse_sheet_cells(root, shared)


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    """Izvuci vrijednost celije prema tipu (t): inlineStr / s / str / b / broj."""
    ctype = cell.get("t")
    if ctype == "inlineStr":
        for child in cell:
            if _local(child.tag) == "is":
                return _collect_text(child)
        return ""
    # <v> tekst (shared index, formula rezultat ili broj)
    vtext = ""
    for child in cell:
        if _local(child.tag) == "v":
            vtext = child.text or ""
            break
    if ctype == "s":
        try:
            return shared[int(vtext)]
        except (ValueError, IndexError):
            return ""
    if ctype in ("str", "e"):     # formula-string / greska: doslovno
        return vtext
    if ctype == "b":
        return "TRUE" if vtext.strip() == "1" else "FALSE"
    # broj (t izostavljen ili "n")
    return _normalize_number(vtext) if vtext != "" else ""


def _parse_sheet_cells(root: ET.Element, shared: list[str]) -> list[list[str]]:
    """Parsiraj sheetData u 2D listu (rijetke celije popunjene praznim)."""
    cells: dict[tuple[int, int], str] = {}   # (row0, col0) -> value
    max_row = 0
    max_col = 0
    for sheet_data in root:
        if _local(sheet_data.tag) != "sheetData":
            continue
        for row in sheet_data:
            if _local(row.tag) != "row":
                continue
            row_attr = row.get("r")
            row_num = int(row_attr) if row_attr else None
            col_cursor = 0
            for cell in row:
                if _local(cell.tag) != "c":
                    continue
                ref = cell.get("r")
                if ref:
                    col0, r1 = _split_ref(ref)
                else:
                    col0 = col_cursor
                    r1 = row_num if row_num else 1
                col_cursor = col0 + 1
                val = _cell_value(cell, shared)
                cells[(r1 - 1, col0)] = val
                if r1 > max_row:
                    max_row = r1
                if col0 + 1 > max_col:
                    max_col = col0 + 1
    grid: list[list[str]] = []
    for r in range(max_row):
        grid.append([cells.get((r, c), "") for c in range(max_col)])
    return grid


# ---------------------------------------------------------------------------
# write_column_a
# ---------------------------------------------------------------------------

def write_column_a(path: str, sheet: str, values: list[str]) -> list[str]:
    """Upisi values u stupac A, redci 1..len(values) zadanog sheeta. NE brise
    sadrzaj ispod retka N. Vraca listu upozorenja. Ako datoteka/sheet ne postoji,
    kreira ih. Atomican zapis uz 3 pokusaja os.replace (za slucaj zakljucane
    datoteke)."""
    _check_zip(path)
    warnings: list[str] = []

    if not os.path.isfile(path):
        new_bytes = _build_new_workbook(sheet, values)
        _atomic_write(path, new_bytes)
        return warnings

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
        part = _find_sheet_part(zf, sheet)

    if part is None:
        # datoteka postoji, ali sheet ne — dodaj novi sheet part
        parts = _add_sheet(parts, sheet, values)
    else:
        new_sheet_xml, warnings = _write_sheet_column_a(parts[part], values)
        parts[part] = new_sheet_xml

    out_bytes = _repack_zip(parts)
    _atomic_write(path, out_bytes)
    return warnings


def _write_sheet_column_a(sheet_bytes: bytes, values: list[str]) -> tuple[bytes, list[str]]:
    """Izmijeni postojeci worksheet XML: postavi stupac A (redci 1..N) na
    inlineStr vrijednosti, ostale celije/retke ostavi. Vrati (novi_xml, upozorenja)."""
    ET.register_namespace("", _NS_MAIN)
    ET.register_namespace("r", _NS_REL)
    root = ET.fromstring(sheet_bytes)
    main = f"{{{_NS_MAIN}}}"

    sheet_data = None
    for child in root:
        if _local(child.tag) == "sheetData":
            sheet_data = child
            break
    if sheet_data is None:
        sheet_data = ET.SubElement(root, main + "sheetData")

    # postojeci redci po broju
    rows_by_num: dict[int, ET.Element] = {}
    for row in list(sheet_data):
        if _local(row.tag) != "row":
            continue
        r_attr = row.get("r")
        if r_attr:
            rows_by_num[int(r_attr)] = row

    warnings: list[str] = []
    n = len(values)
    # upozorenje: ne-prazne celije u stupcu A ispod retka N
    for rnum, row in rows_by_num.items():
        if rnum <= n:
            continue
        for cell in row:
            if _local(cell.tag) != "c":
                continue
            ref = cell.get("r")
            if not ref:
                continue
            col0, _ = _split_ref(ref)
            if col0 == 0:
                val = _cell_value(cell, [])
                if val.strip():
                    warnings.append(
                        f"stupac A, redak {rnum} ima sadrzaj '{val}' ispod "
                        f"zadnjeg upisanog retka {n} (nije obrisano)")

    for i, value in enumerate(values):
        rnum = i + 1
        row = rows_by_num.get(rnum)
        if row is None:
            row = ET.Element(main + "row", {"r": str(rnum)})
            rows_by_num[rnum] = row
        _set_cell_inline(row, 0, rnum, value, main)

    # posloziti retke u sheetData po broju, celije po stupcu
    for row in list(sheet_data):
        sheet_data.remove(row)
    for rnum in sorted(rows_by_num):
        row = rows_by_num[rnum]
        _sort_row_cells(row, main)
        row.set("r", str(rnum))
        # spans/dimension prepusti Excelu; ne diramo ostale atribute retka
        sheet_data.append(row)

    _update_dimension(root, sheet_data, main)

    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return xml, warnings


def _set_cell_inline(row: ET.Element, col0: int, rnum: int, value: str, main: str) -> None:
    """Postavi celiju (col0, rnum) na inlineStr vrijednost; sacuvaj stil (s)."""
    ref = _index_to_col(col0) + str(rnum)
    target = None
    for cell in row:
        if _local(cell.tag) == "c" and cell.get("r") == ref:
            target = cell
            break
    if target is None:
        target = ET.SubElement(row, main + "c", {"r": ref})
    # ocisti sadrzaj (v/f/is), zadrzi eventualni stil s
    style = target.get("s")
    for child in list(target):
        target.remove(child)
    target.attrib.clear()
    target.set("r", ref)
    if style is not None:
        target.set("s", style)
    target.set("t", "inlineStr")
    is_el = ET.SubElement(target, main + "is")
    t_el = ET.SubElement(is_el, main + "t")
    if value != value.strip() or value == "":
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = value


def _sort_row_cells(row: ET.Element, main: str) -> None:
    cells = [c for c in row if _local(c.tag) == "c"]
    others = [c for c in row if _local(c.tag) != "c"]
    cells.sort(key=lambda c: _split_ref(c.get("r"))[0] if c.get("r") else 0)
    for c in list(row):
        row.remove(c)
    for c in others:
        row.append(c)
    for c in cells:
        row.append(c)


def _update_dimension(root: ET.Element, sheet_data: ET.Element, main: str) -> None:
    """Osvjezi <dimension ref="A1:..."> ako postoji (Excel inace zna zatraziti
    'repair'). Racuna max redak i max stupac iz svih celija."""
    dim = None
    for child in root:
        if _local(child.tag) == "dimension":
            dim = child
            break
    if dim is None:
        return
    max_row = 0
    max_col = 0
    for row in sheet_data:
        if _local(row.tag) != "row":
            continue
        for cell in row:
            if _local(cell.tag) != "c":
                continue
            ref = cell.get("r")
            if not ref:
                continue
            col0, r1 = _split_ref(ref)
            max_row = max(max_row, r1)
            max_col = max(max_col, col0 + 1)
    if max_row == 0:
        return
    dim.set("ref", f"A1:{_index_to_col(max_col - 1)}{max_row}")


# ---------------------------------------------------------------------------
# kreiranje novog workbooka / dodavanje sheeta
# ---------------------------------------------------------------------------

def _sheet_xml_from_values(values: list[str]) -> bytes:
    """Generiraj cist worksheet XML sa stupcem A = values (inlineStr)."""
    rows_xml = []
    for i, value in enumerate(values):
        rnum = i + 1
        safe = escape(value)
        rows_xml.append(
            f'<row r="{rnum}"><c r="A{rnum}" t="inlineStr">'
            f'<is><t xml:space="preserve">{safe}</t></is></c></row>')
    last = len(values) if values else 1
    body = "".join(rows_xml)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
        f'<dimension ref="A1:A{last}"/>'
        f'<sheetData>{body}</sheetData>'
        '</worksheet>'
    ).encode("utf-8")


def _build_new_workbook(sheet: str, values: list[str]) -> bytes:
    """Kreiraj minimalni, ali valjani .xlsx s jednim sheetom (values u stupcu A)."""
    safe_sheet = escape(sheet, {'"': "&quot;"})
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Types xmlns="{_NS_CT}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>'
        ).encode("utf-8"),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{_NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ).encode("utf-8"),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
            f'<sheets><sheet name="{safe_sheet}" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ).encode("utf-8"),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{_NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ).encode("utf-8"),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<styleSheet xmlns="{_NS_MAIN}">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf/></cellXfs>'
            '</styleSheet>'
        ).encode("utf-8"),
        "xl/worksheets/sheet1.xml": _sheet_xml_from_values(values),
    }
    return _repack_zip(parts)


def _add_sheet(parts: dict[str, bytes], sheet: str, values: list[str]) -> dict[str, bytes]:
    """Dodaj novi sheet u postojeci workbook (workbook.xml + rels + Content_Types
    + novi sheet part). Vrati azuriran parts dict."""
    parts = dict(parts)
    ET.register_namespace("", _NS_MAIN)
    ET.register_namespace("r", _NS_REL)

    # 1) odaberi slobodno ime sheet parta i rId
    existing_sheet_parts = [n for n in parts if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
    idx = 1
    while f"xl/worksheets/sheet{idx}.xml" in parts:
        idx += 1
    new_part = f"xl/worksheets/sheet{idx}.xml"

    rels_name = "xl/_rels/workbook.xml.rels"
    rels_root = ET.fromstring(parts[rels_name]) if rels_name in parts else \
        ET.fromstring(f'<Relationships xmlns="{_NS_PKGREL}"></Relationships>')
    used_ids = {rel.get("Id") for rel in rels_root}
    rid_num = 1
    while f"rId{rid_num}" in used_ids:
        rid_num += 1
    new_rid = f"rId{rid_num}"
    rel = ET.SubElement(rels_root, f"{{{_NS_PKGREL}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
    rel.set("Target", f"worksheets/sheet{idx}.xml")
    parts[rels_name] = ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)

    # 2) workbook.xml — dodaj <sheet>
    wb_root = ET.fromstring(parts["xl/workbook.xml"])
    main = f"{{{_NS_MAIN}}}"
    sheets_el = None
    for child in wb_root:
        if _local(child.tag) == "sheets":
            sheets_el = child
            break
    if sheets_el is None:
        sheets_el = ET.SubElement(wb_root, main + "sheets")
    max_sheet_id = 0
    for s in sheets_el:
        try:
            max_sheet_id = max(max_sheet_id, int(s.get("sheetId", "0")))
        except ValueError:
            pass
    new_sheet_el = ET.SubElement(sheets_el, main + "sheet")
    new_sheet_el.set("name", sheet)
    new_sheet_el.set("sheetId", str(max_sheet_id + 1))
    new_sheet_el.set(f"{{{_NS_REL}}}id", new_rid)
    parts["xl/workbook.xml"] = ET.tostring(wb_root, encoding="UTF-8", xml_declaration=True)

    # 3) [Content_Types].xml — Override za novi part
    ct_name = "[Content_Types].xml"
    ct_root = ET.fromstring(parts[ct_name])
    override = ET.SubElement(ct_root, f"{{{_NS_CT}}}Override")
    override.set("PartName", f"/{new_part}")
    override.set("ContentType",
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    parts[ct_name] = ET.tostring(ct_root, encoding="UTF-8", xml_declaration=True)

    # 4) sam sheet part
    parts[new_part] = _sheet_xml_from_values(values)
    _ = existing_sheet_parts  # (samo za citljivost — brojanje iznad)
    return parts


# ---------------------------------------------------------------------------
# ZIP repack + atomican zapis
# ---------------------------------------------------------------------------

def _repack_zip(parts: dict[str, bytes]) -> bytes:
    import io
    buf = io.BytesIO()
    # [Content_Types].xml prvi (konvencija), ostalo redom
    order = sorted(parts, key=lambda n: (n != "[Content_Types].xml", n))
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            zf.writestr(name, parts[name])
    return buf.getvalue()


def _atomic_write(path: str, data: bytes, attempts: int = 3) -> None:
    """Zapisi u temp u istom direktoriju pa os.replace (atomicno). Do `attempts`
    pokusaja zbog moguce zakljucane datoteke (otvorene u Excelu)."""
    directory = os.path.dirname(os.path.abspath(path))
    tmp = os.path.join(directory, f".~xlsx_lite_{os.getpid()}_{int(time.time()*1000)}.tmp")
    with open(tmp, "wb") as fh:
        fh.write(data)
    last_err: Exception | None = None
    for i in range(attempts):
        try:
            os.replace(tmp, path)
            return
        except OSError as exc:
            last_err = exc
            time.sleep(0.5 * (i + 1))
    # neuspjeh — pocisti temp i podigni gresku
    try:
        os.remove(tmp)
    except OSError:
        pass
    raise OSError(
        f"ne mogu zapisati '{path}' nakon {attempts} pokusaja "
        f"(je li otvorena u Excelu?): {last_err}")
