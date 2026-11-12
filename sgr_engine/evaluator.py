"""Minimal, safe evaluator for the arithmetic formulas used in station sheets.

openpyxl never computes formulas, and a workbook saved by openpyxl carries no
cached results.  The rollover needs *values* of yesterday's formula cells
(e.g. ``AVAILABLE CASH`` -> today's ``ADD UP``), so we evaluate them here.

Supported: numbers, same-sheet cell refs (``$`` allowed), ranges inside
functions, ``+ - * / ^``, unary minus, parentheses, ``%`` suffix, and the
functions SUM, ROUND, ABS, MIN, MAX, AVERAGE.  Anything else raises
``FormulaError`` so the caller can fall back to a cached value.
"""
from __future__ import annotations

import math
import re
from typing import Callable, Dict, List, Optional

from openpyxl.utils import column_index_from_string, get_column_letter

__all__ = ["FormulaError", "SheetEvaluator"]


class FormulaError(Exception):
    pass


_TOKEN = re.compile(
    r"""\s*(?:
        (?P<range>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+) |
        (?P<func>[A-Z][A-Z0-9.]*)\s*\( |
        (?P<ref>\$?[A-Z]{1,3}\$?\d+) |
        (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+) |
        (?P<op>[-+*/^(),%])
    )""",
    re.X,
)

_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")


def _split_ref(ref: str):
    m = _REF.fullmatch(ref)
    if not m:
        raise FormulaError(f"bad ref {ref}")
    return column_index_from_string(m.group(1)), int(m.group(2))


def _to_num(v) -> float:
    if v is None or isinstance(v, str):
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    raise FormulaError(f"non numeric {v!r}")


_FUNCS: Dict[str, Callable] = {
    "SUM": lambda *a: sum(_flat(a)),
    "MIN": lambda *a: min(_flat(a)) if _flat(a) else 0.0,
    "MAX": lambda *a: max(_flat(a)) if _flat(a) else 0.0,
    "AVERAGE": lambda *a: (sum(_flat(a)) / len(_flat(a))) if _flat(a) else 0.0,
    "ABS": lambda x: abs(_scalar(x)),
    "ROUND": lambda x, n=0: _xl_round(_scalar(x), int(_scalar(n))),
}


def _xl_round(x: float, n: int) -> float:
    f = 10.0 ** n
    return math.floor(abs(x) * f + 0.5) / f * (1 if x >= 0 else -1)


def _flat(args) -> List[float]:
    out: List[float] = []
    for a in args:
        if isinstance(a, list):
            out.extend(a)
        else:
            out.append(a)
    return out


def _scalar(x) -> float:
    if isinstance(x, list):
        return sum(x)
    return x


class _Parser:
    """Recursive-descent parser producing a value directly."""

    def __init__(self, text: str, resolve: Callable[[int, int], float]):
        self.toks = []
        pos = 0
        text = text.upper()
        while pos < len(text):
            if text[pos].isspace():
                pos += 1
                continue
            m = _TOKEN.match(text, pos)
            if not m or m.end() == pos:
                raise FormulaError(f"cannot tokenize at {text[pos:]!r}")
            kind = m.lastgroup
            self.toks.append((kind, m.group(kind)))
            pos = m.end()
        self.i = 0
        self.resolve = resolve

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, val=None):
        tok = self.peek()
        if val is not None and tok[1] != val:
            raise FormulaError(f"expected {val} got {tok}")
        self.i += 1
        return tok

    def parse(self):
        v = self.expr()
        if self.i != len(self.toks):
            raise FormulaError(f"trailing tokens {self.toks[self.i:]}")
        return _scalar(v)

    def expr(self):
        v = self.term()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            r = self.term()
            v = _scalar(v) + _scalar(r) if op == "+" else _scalar(v) - _scalar(r)
        return v

    def term(self):
        v = self.power()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            r = _scalar(self.power())
            if op == "*":
                v = _scalar(v) * r
            else:
                if r == 0:
                    raise FormulaError("#DIV/0!")
                v = _scalar(v) / r
        return v

    def power(self):
        v = self.unary()
        while self.peek()[1] == "^":
            self.take()
            v = _scalar(v) ** _scalar(self.unary())
        return v

    def unary(self):
        if self.peek()[1] == "-":
            self.take()
            return -_scalar(self.unary())
        if self.peek()[1] == "+":
            self.take()
            return self.unary()
        v = self.atom()
        while self.peek()[1] == "%":
            self.take()
            v = _scalar(v) / 100.0
        return v

    def atom(self):
        kind, val = self.take()
        if kind == "num":
            return float(val)
        if kind == "ref":
            c, r = _split_ref(val)
            return self.resolve(c, r)
        if kind == "range":
            a, b = val.split(":")
            c1, r1 = _split_ref(a)
            c2, r2 = _split_ref(b)
            return [
                self.resolve(c, r)
                for r in range(min(r1, r2), max(r1, r2) + 1)
                for c in range(min(c1, c2), max(c1, c2) + 1)
            ]
        if kind == "func":
            fn = _FUNCS.get(val)
            if fn is None:
                raise FormulaError(f"unsupported function {val}")
            args = []
            if self.peek()[1] != ")":
                args.append(self.expr())
                while self.peek()[1] == ",":
                    self.take()
                    args.append(self.expr())
            self.take(")")
            return fn(*args)
        if val == "(":
            v = self.expr()
            self.take(")")
            return v
        raise FormulaError(f"unexpected token {val!r}")


class SheetEvaluator:
    """Evaluate cells of one worksheet, with memoisation and cycle detection.

    ``cached`` is an optional worksheet loaded with ``data_only=True``; its
    values are used as a fallback when a formula cannot be evaluated.
    """

    def __init__(self, ws, cached=None):
        self.ws = ws
        self.cached = cached
        self._memo: Dict[tuple, float] = {}
        self._stack: set = set()

    def invalidate(self):
        self._memo.clear()

    def value(self, col: int, row: int) -> float:
        key = (col, row)
        if key in self._memo:
            return self._memo[key]
        raw = self.ws.cell(row=row, column=col).value
        if isinstance(raw, str) and raw.startswith("="):
            if key in self._stack:
                raise FormulaError(f"circular ref at {get_column_letter(col)}{row}")
            self._stack.add(key)
            try:
                v = _Parser(raw[1:], self.value).parse()
            except FormulaError:
                v = self._fallback(col, row)
                if v is None:
                    raise
            finally:
                self._stack.discard(key)
        else:
            v = _to_num(raw)
        self._memo[key] = v
        return v

    def value_at(self, coord: str) -> float:
        c, r = _split_ref(coord)
        return self.value(c, r)

    def try_value(self, col: int, row: int) -> Optional[float]:
        try:
            return self.value(col, row)
        except FormulaError:
            return None

    def _fallback(self, col, row):
        if self.cached is None:
            return None
        v = self.cached.cell(row=row, column=col).value
        return float(v) if isinstance(v, (int, float)) else None
