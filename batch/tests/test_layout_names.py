#!/usr/bin/env python3
"""Testovi za layout_names.py — primjeri iz LISP komentara (ele:nums,
ele:valid-layout-p, ele:layout<)."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import layout_names as ln  # noqa: E402


class TestNums(unittest.TestCase):
    def test_primjeri_iz_lisp_komentara(self):
        # ";; Extracts the numbers from a string, e.g. "10-2x3" -> (10 2 3)"
        self.assertEqual(ln.nums("10-2x3"), [10, 2, 3])
        self.assertEqual(ln.nums("3-1x1"), [3, 1, 1])

    def test_minus_izmedu_znamenki_nije_predznak(self):
        # U "10-2x3" crtica je izmedu znamenki -> razdvaja, ne predznak.
        self.assertEqual(ln.nums("1-2x3"), [1, 2, 3])

    def test_vodeci_minus(self):
        # Vodeci minus (nije iza znamenke) se cuva kao predznak.
        self.assertEqual(ln.nums("-5x3"), [-5, 3])

    def test_decimalna_tocka(self):
        self.assertEqual(ln.nums("1.5x2"), [1.5, 2])

    def test_znamenka_u_slovu(self):
        # "A4L" nije valjano ime, ali ele:nums svejedno izvuce znamenku 4.
        self.assertEqual(ln.nums("A4L"), [4])

    def test_bez_brojeva(self):
        self.assertEqual(ln.nums("ABC"), [])


class TestValidLayout(unittest.TestCase):
    def test_prihvaca(self):
        for name in ("3-1x1", "10-2x3", "1-1x1", "100-200x300"):
            self.assertTrue(ln.valid_layout(name), name)

    def test_odbacuje(self):
        for name in ("1-2-3", "A4L", "1x1", "3-1x1x2", "", "3-x1",
                     "-1x1", "3-1x", "3--1x1", "3-1X1", " 3-1x1", "3-1x1 "):
            self.assertFalse(ln.valid_layout(name), name)


class TestSort(unittest.TestCase):
    def test_layout_sort_key(self):
        self.assertEqual(ln.layout_sort_key("10-2x3"), (10, 2))
        self.assertEqual(ln.layout_sort_key("3-1x1"), (3, 1))

    def test_sort_po_prvom_pa_drugom(self):
        data = ["10-1x1", "2-5x1", "2-1x1", "1-9x1"]
        self.assertEqual(ln.sort_valid_layouts(data),
                         ["1-9x1", "2-1x1", "2-5x1", "10-1x1"])

    def test_treci_broj_se_ignorira(self):
        # Isti (prvi, drugi) -> stabilan poredak (ulazni redoslijed ostaje).
        data = ["5-1x9", "5-1x2", "5-1x7"]
        self.assertEqual(ln.sort_valid_layouts(data), ["5-1x9", "5-1x2", "5-1x7"])

    def test_filter_izbacuje_nevaljane(self):
        data = ["3-1x1", "A4L", "2-1x1", "CUSTOM", "1-1x1"]
        self.assertEqual(ln.sort_valid_layouts(data), ["1-1x1", "2-1x1", "3-1x1"])


if __name__ == "__main__":
    unittest.main()
