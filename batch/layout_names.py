#!/usr/bin/env python3
"""
layout_names.py — port ele:nums / ele:valid-layout-p / ele:layout< iz
AutoLisp/ExportLayoutsToExcel.lsp.

Tri funkcije koje kopiraju semantiku LISP-a:
  - nums(str)            -> lista brojeva iz stringa  ("10-2x3" -> [10, 2, 3])
  - valid_layout(str)    -> True samo za uzorak  ^[0-9]+-[0-9]+x[0-9]+$
  - layout_sort_key(str) -> kljuc za stabilno sortiranje (prvi pa drugi broj)

VAZNO (machine-parni port): ova logika MORA ostati identicna C# modulu
LayoutCrator/LayoutCreatorCore/LayoutNameOrder.cs. Mijenjaj ih SAMO zajedno —
oba su port istih ele: funkcija i koriste ih i sortiranje tabova i export.

SAMO stdlib.
"""

from __future__ import annotations

import re

# Uzorak valjanog imena layouta: znamenke, JEDNA crtica, znamenke, JEDNO malo
# 'x', znamenke. Vjeran port ele:valid-layout-p (vl-string-search je case-
# sensitive pa je 'x' iskljucivo malo; ele:all-digits-p je striktno ASCII 0-9,
# zato [0-9] a ne \d).
_VALID_RE = re.compile(r"^[0-9]+-[0-9]+x[0-9]+$")


def _is_digit(code: int | None) -> bool:
    """Kao ele:digit-p: (< 47 c 58) — striktno ASCII 0-9."""
    return code is not None and 48 <= code <= 57


def _nums_char(a: int | None, b: int, c: int | None) -> int:
    """Port ele:nums-char. a = kod prethodnog znaka, b = tekuci, c = sljedeci
    (a/c mogu biti None na rubovima). Zadrzi b ako je znamenka, vodeci minus,
    ili decimalna tocka izmedu dvije znamenke; inace zamijeni razmakom (32)."""
    if 48 <= b <= 57:
        return b
    if b == 45 and _is_digit(c) and not _is_digit(a):        # vodeci '-'
        return b
    if b == 46 and _is_digit(a) and _is_digit(c):            # '.' izmedu znamenki
        return b
    return 32


def _parse_token(tok: str) -> int | float:
    """LISP 'read' na tokenu daje int za cijele brojeve, float za decimalne."""
    try:
        return int(tok)
    except ValueError:
        return float(tok)


def nums(text: str) -> list[int | float]:
    """Izvlaci brojeve iz stringa, npr. "10-2x3" -> [10, 2, 3].
    Port ele:nums (Lee Mac numbersFromString)."""
    codes = [ord(ch) for ch in text]
    prev = [None] + codes[:-1]
    nxt = codes[1:] + [None]
    filtered = "".join(chr(_nums_char(a, b, c)) for a, b, c in zip(prev, codes, nxt))
    return [_parse_token(tok) for tok in filtered.split()]


def valid_layout(name: str) -> bool:
    """True samo za uzorak ^[0-9]+-[0-9]+x[0-9]+$  (port ele:valid-layout-p).
    Primjeri: "3-1x1", "10-2x3" -> True; "1-2-3", "A4L", "1x1", "3-1x1x2" -> False."""
    return bool(_VALID_RE.match(name))


def layout_sort_key(name: str) -> tuple[int | float, int | float]:
    """Kljuc za stabilno sortiranje: (prvi broj, drugi broj). Treci se ignorira —
    vjerno ele:layout< koji usporeduje samo (car) pa (cadr). Koristi se s
    Pythonovim stabilnim sortom (list.sort / sorted)."""
    n = nums(name)
    first = n[0] if len(n) > 0 else 0
    second = n[1] if len(n) > 1 else 0
    return (first, second)


def sort_valid_layouts(names: list[str]) -> list[str]:
    """Filtriraj na valjana imena i stabilno sortiraj po layout_sort_key.
    Ekvivalent (vl-sort (vl-remove-if-not 'ele:valid-layout-p ...) 'ele:layout<),
    ali sa STABILNIM sortom (plan §4.3) umjesto vl-sort semantike."""
    valid = [n for n in names if valid_layout(n)]
    return sorted(valid, key=layout_sort_key)
