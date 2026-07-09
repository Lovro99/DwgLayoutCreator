#!/usr/bin/env python3
"""Testovi za xlsx_lite.py — round-trip (modul sam generira minimalni xlsx pa
ga cita natrag), citanje sharedStrings i inlineStr, ponasanje write_column_a
(cuvanje ostalih stupaca, upozorenje ispod retka N, dodavanje sheeta, novi file)."""

import io
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import xlsx_lite as xl  # noqa: E402

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _minimal_xlsx_shared(cells_rows):
    """Sagradi xlsx s sharedStrings za dane retke (lista listi vrijednosti).
    Sve vrijednosti tretirane kao shared stringovi."""
    # sabrati jedinstvene stringove
    uniq = []
    index = {}
    for row in cells_rows:
        for v in row:
            if v not in index:
                index[v] = len(uniq)
                uniq.append(v)
    sst = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<sst xmlns="{NS_MAIN}" count="{sum(len(r) for r in cells_rows)}" '
           f'uniqueCount="{len(uniq)}">'
           + "".join(f"<si><t>{v}</t></si>" for v in uniq) + "</sst>")
    rows_xml = []
    for r, row in enumerate(cells_rows, start=1):
        cells = []
        for c, v in enumerate(row):
            ref = xl._index_to_col(c) + str(r)
            cells.append(f'<c r="{ref}" t="s"><v>{index[v]}</v></c>')
        rows_xml.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<worksheet xmlns="{NS_MAIN}"><sheetData>'
             + "".join(rows_xml) + "</sheetData></worksheet>")
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{NS_CT}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>').encode(),
        "_rels/.rels": (
            f'<Relationships xmlns="{NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>').encode(),
        "xl/workbook.xml": (
            f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">'
            '<sheets><sheet name="Podaci" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>').encode(),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>').encode(),
        "xl/sharedStrings.xml": sst.encode(),
        "xl/worksheets/sheet1.xml": sheet.encode(),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _xlsx_with_numbers_and_inline(path):
    """xlsx s brojevima (<v> bez t) i inlineStr celijama u sheetu "Podaci"."""
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<worksheet xmlns="{NS_MAIN}"><sheetData>'
             '<row r="1"><c r="A1" t="inlineStr"><is><t>INVESTITOR</t></is></c>'
             '<c r="B1" t="inlineStr"><is><t>HEP ODS</t></is></c></row>'
             '<row r="2"><c r="A2" t="inlineStr"><is><t>BROJ</t></is></c>'
             '<c r="B2"><v>42</v></c></row>'
             '<row r="3"><c r="A3" t="inlineStr"><is><t>FAKTOR</t></is></c>'
             '<c r="B3"><v>1.5</v></c></row>'
             '<row r="4"><c r="A4" t="inlineStr"><is><t>CIJELI</t></is></c>'
             '<c r="B4"><v>42.0</v></c></row>'
             '</sheetData></worksheet>')
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{NS_CT}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>').encode(),
        "_rels/.rels": (
            f'<Relationships xmlns="{NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>').encode(),
        "xl/workbook.xml": (
            f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">'
            '<sheets><sheet name="Podaci" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>').encode(),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{NS_PKGREL}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '</Relationships>').encode(),
        "xl/worksheets/sheet1.xml": sheet.encode(),
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)


class TestReadShared(unittest.TestCase):
    def test_shared_strings(self):
        data = _minimal_xlsx_shared([["INVESTITOR", "HEP ODS"], ["GRAD", "TS 10/0.4"]])
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "podaci.xlsx")
            with open(p, "wb") as fh:
                fh.write(data)
            grid = xl.read_sheet(p, "Podaci")
        self.assertEqual(grid, [["INVESTITOR", "HEP ODS"], ["GRAD", "TS 10/0.4"]])

    def test_case_insensitive_sheet(self):
        data = _minimal_xlsx_shared([["A", "B"]])
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "x.xlsx")
            with open(p, "wb") as fh:
                fh.write(data)
            grid = xl.read_sheet(p, "podaci")   # razlicit case
        self.assertEqual(grid, [["A", "B"]])


class TestReadNumbersInline(unittest.TestCase):
    def test_inline_i_normalizacija_brojeva(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "num.xlsx")
            _xlsx_with_numbers_and_inline(p)
            grid = xl.read_sheet(p, "Podaci")
        self.assertEqual(grid[0], ["INVESTITOR", "HEP ODS"])
        self.assertEqual(grid[1], ["BROJ", "42"])        # cijeli broj bez .0
        self.assertEqual(grid[2], ["FAKTOR", "1.5"])
        self.assertEqual(grid[3], ["CIJELI", "42"])      # 42.0 -> "42"


class TestRoundTrip(unittest.TestCase):
    def test_write_new_file_then_read(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "nacrti.xlsx")
            warns = xl.write_column_a(p, "Nacrti", ["1-1x1", "2-1x1", "10-2x3"])
            self.assertEqual(warns, [])
            self.assertTrue(os.path.isfile(p))
            grid = xl.read_sheet(p, "Nacrti")
        self.assertEqual([r[0] for r in grid], ["1-1x1", "2-1x1", "10-2x3"])

    def test_dodavanje_sheeta_u_postojeci(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "book.xlsx")
            xl.write_column_a(p, "Podaci", ["INVESTITOR", "GRAD"])
            xl.write_column_a(p, "Nacrti", ["1-1x1", "2-1x1"])
            # oba sheeta procitljiva
            self.assertEqual([r[0] for r in xl.read_sheet(p, "Podaci")],
                             ["INVESTITOR", "GRAD"])
            self.assertEqual([r[0] for r in xl.read_sheet(p, "Nacrti")],
                             ["1-1x1", "2-1x1"])

    def test_cuva_ostale_stupce_ne_brise_ispod(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "num.xlsx")
            _xlsx_with_numbers_and_inline(p)   # ima B stupac + 4 retka
            # upisi 2 vrijednosti u A (redci 1-2), ispod (redci 3-4) A je prazan
            warns = xl.write_column_a(p, "Podaci", ["X1", "X2"])
            grid = xl.read_sheet(p, "Podaci")
        # stupac A prepisan u prva 2 retka
        self.assertEqual(grid[0][0], "X1")
        self.assertEqual(grid[1][0], "X2")
        # stupac B ocuvan (nije obrisan)
        self.assertEqual(grid[0][1], "HEP ODS")
        self.assertEqual(grid[1][1], "42")
        # redci 3-4 (ispod N=2) netaknuti u stupcu B
        self.assertEqual(grid[2][1], "1.5")
        self.assertEqual(grid[3][1], "42")
        # stupac A redci 3-4 imali su FAKTOR/CIJELI -> 2 upozorenja (nisu obrisani)
        self.assertEqual(len(warns), 2)
        self.assertEqual(grid[2][0], "FAKTOR")
        self.assertEqual(grid[3][0], "CIJELI")

    def test_upozorenje_ispod_retka_n(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "nacrti.xlsx")
            xl.write_column_a(p, "Nacrti", ["1-1x1", "2-1x1", "3-1x1", "4-1x1"])
            # sada upisi samo 2 -> redci 3,4 u stupcu A imaju stari sadrzaj
            warns = xl.write_column_a(p, "Nacrti", ["A", "B"])
        self.assertEqual(len(warns), 2)
        self.assertIn("redak 3", warns[0])
        self.assertIn("redak 4", warns[1])


class TestXlsGuard(unittest.TestCase):
    def test_xls_nije_podrzan(self):
        with self.assertRaises(ValueError):
            xl.read_sheet("/tmp/nepostoji.xls", "Podaci")
        with self.assertRaises(ValueError):
            xl.write_column_a("/tmp/nepostoji.xls", "Podaci", ["a"])


if __name__ == "__main__":
    unittest.main()
