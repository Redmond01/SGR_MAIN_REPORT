"""Anchor-based layout detection for SGR daily station sheets.

Nothing here relies on fixed coordinates.  Every table is located from its
keyword headers (``OPENING METER``, ``TANK``/``OPENNG DIPPG``, ``ADD UP`` ...)
and its extent is derived from the cells around it.

A sheet is two mirror halves: LEFT (outer / station meters) and RIGHT
(inner meters / totalizers).  Each half is parsed into a :class:`HalfLayout`.
"""
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

__all__ = [
    "LayoutError",
    "Pump",
    "SalesTable",
    "Tank",
    "TankTable",
    "LabeledValues",
    "ListBlock",
    "PosTable",
    "HalfLayout",
    "SheetLayout",
    "parse_sheet",
    "norm",
    "product_key",
    "is_formula",
]


class LayoutError(Exception):
    pass


# ---------------------------------------------------------------- helpers

def norm(v) -> str:
    """Upper-case, collapse whitespace.  Used for every anchor comparison."""
    return re.sub(r"\s+", " ", str(v)).strip().upper() if v is not None else ""


def compact(v) -> str:
    """Label key without spaces/punctuation: 'PMS 1' == 'PMS1' == 'pms-1'."""
    return re.sub(r"[^A-Z0-9]", "", norm(v))


def product_key(label) -> str:
    """Product family of a pump/tank label: PMS, AGO, DPK, LPG (or '')."""
    n = compact(label)
    for p in ("PMS", "AGO", "DPK", "LPG", "HHK"):
        if n.startswith(p):
            return "DPK" if p == "HHK" else p
    return ""


def is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def coord(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"


# Header synonyms (normalised).  Station sheets use these exact misspellings.
OPENING_METER = {"OPENING METER", "OPENING METERS", "OPENING"}
CLOSING_METER = {"CLOSING METER", "CLOSING METERS", "CLOSING"}
OPEN_DIP = {
    "OPENNG DIPPG",
    "OPENING DIPPING",
    "OPENING DIP",
    "OPENNG DIPPING",
    "OPENING DIPPG",
    "OPEN DIPPG",
    "OPEN DIPPING",
}
CLOSE_DIP = {"CLOSNG DIPPG", "CLOSING DIPPING", "CLOSING DIP", "CLOSNG DIPPING", "CLOSING DIPPG"}
TANK = {"TANK", "TANKS"}
ADD_UP = {"ADD UP", "ADDUP", "ADD-UP"}
AVAILABLE_CASH = {"AVAILABLE CASH", "AVAILABLE  CASH"}
PRODUCTS_RECEIVED = {"PRODUCTS RECEIVED", "PRODUCT RECEIVED", "PRODUCTS RECIEVED"}
EXPENSES = {"EXPENSES", "EXPENSE"}
IMPREST = {"IMPREST"}
BANK_LODGMENT = {"BANK LODGMENT", "BANK LODGEMENT"}


# ---------------------------------------------------------------- model

@dataclass
class Pump:
    label: str
    row: int
    product: str


@dataclass
class SalesTable:
    header_row: int
    label_col: int
    open_col: int
    closing_cols: List[int]          # usually 4: C1..C4
    sales_cols: List[int]            # SALES 1..4 (formulas)
    price_cols: List[int]            # PRICE 1..4
    amount_col: Optional[int]
    pumps: List[Pump] = field(default_factory=list)
    total_rows: List[int] = field(default_factory=list)

    def pump(self, label) -> Optional[Pump]:
        k = compact(label)
        return next((p for p in self.pumps if compact(p.label) == k), None)


@dataclass
class Tank:
    label: str
    row: int
    product: str
    # list of (open_col, close_col) pairs; pair 0 is the left-most
    linked_into: Optional[str] = None   # label of host tank when manifolded
    hosts: List[str] = field(default_factory=list)  # labels merged into this tank
    has_sales_formula: bool = True
    received_row: Optional[int] = None  # linked row in PRODUCTS RECEIVED if referenced in formula


@dataclass
class TankTable:
    header_row: int
    label_col: int
    dip_pairs: List[Tuple[int, int]]   # (opening col, closing col)
    sales_col: Optional[int]
    gear_cols: List[int]
    tanks: List[Tank] = field(default_factory=list)
    row_limit: int = 10_000

    def tank(self, label) -> Optional[Tank]:
        k = compact(label)
        return next((t for t in self.tanks if compact(t.label) == k), None)


@dataclass
class LabeledValues:
    """A vertical list of label -> value cells (e.g. PRODUCTS RECEIVED)."""
    anchor_row: int
    label_col: int
    value_col: int
    rows: Dict[str, int] = field(default_factory=dict)   # label -> row
    stop_row: Optional[int] = None


@dataclass
class ListBlock:
    """Free-form description/amount list with a SUM total (EXPENSES, IMPREST)."""
    anchor_row: int
    desc_col: int
    amount_col: int
    first_row: int
    last_row: int                    # inclusive, row before the total
    total_row: Optional[int]

    @property
    def capacity(self) -> int:
        return max(0, self.last_row - self.first_row + 1)


@dataclass
class PosTable:
    header_row: int
    value_row: int
    channels: List[Tuple[str, int]]  # (header label, column)
    add_up_col: int
    available_col: Optional[int]
    all_value_cols: List[int] = field(default_factory=list)  # includes blank-header slots to zero on reset


@dataclass
class HalfLayout:
    side: str                         # "LEFT" | "RIGHT"
    col_min: int
    col_max: int
    date_cells: List[str]
    sales: SalesTable
    tanks: Optional[TankTable]
    received: Optional[LabeledValues]
    expenses: Optional[ListBlock]
    imprest: Optional[ListBlock]
    pos: Optional[PosTable]
    bank: Optional[LabeledValues]
    warnings: List[str] = field(default_factory=list)


@dataclass
class SheetLayout:
    title: str
    left: HalfLayout
    right: Optional[HalfLayout]

    @property
    def halves(self) -> List[HalfLayout]:
        return [h for h in (self.left, self.right) if h is not None]


# ---------------------------------------------------------------- scanning

class _Grid:
    """Tiny cache of normalised text for fast anchor lookups."""

    def __init__(self, ws: Worksheet, max_row: int = 120):
        self.ws = ws
        self.max_row = min(ws.max_row, max_row)
        self.max_col = ws.max_column
        self.text: Dict[Tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=1, max_row=self.max_row):
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    self.text[(c.column, c.row)] = norm(c.value)

    def find(self, names, col_min=1, col_max=10_000, row_min=1, row_max=10_000):
        hits = [
            (col, row)
            for (col, row), t in self.text.items()
            if t in names and col_min <= col <= col_max and row_min <= row <= row_max
        ]
        return sorted(hits, key=lambda x: (x[1], x[0]))

    def v(self, col, row):
        return self.ws.cell(row=row, column=col).value

    def t(self, col, row) -> str:
        return self.text.get((col, row), "")


def _parse_sales(g: _Grid, open_col: int, hrow: int, col_max: int, row_stop: int) -> SalesTable:
    closing, sales, price, amount = [], [], [], None
    c = open_col + 1
    while c <= col_max:
        t = g.t(c, hrow)
        if t in CLOSING_METER:
            closing.append(c)
        elif t.startswith("SALES"):
            sales.append(c)
        elif t.startswith("PRICE"):
            price.append(c)
        elif t == "AMOUNT":
            amount = c
            break
        elif t == "" and (sales or price):
            # tolerate one blank header between pairs, stop at second
            if g.t(c + 1, hrow) == "":
                break
        c += 1
    if not closing or not sales:
        raise LayoutError(f"sales header at {coord(open_col, hrow)} incomplete")

    tbl = SalesTable(hrow, open_col - 1, open_col, closing, sales, price, amount)
    for r in range(hrow + 1, row_stop):
        lab = g.v(tbl.label_col, r)
        s1 = g.v(sales[0], r)
        if lab is None or is_formula(lab):
            continue
        n = norm(lab)
        if n.startswith("TOTAL"):
            continue
        if not is_formula(s1):
            continue
        tbl.pumps.append(Pump(str(lab).strip(), r, product_key(lab)))
    for r in range(hrow + 1, row_stop):
        if norm(g.v(tbl.label_col - 1, r)).startswith("TOTAL"):
            tbl.total_rows.append(r)
    if not tbl.pumps:
        raise LayoutError(f"no pump rows under {coord(open_col, hrow)}")
    return tbl


_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")


def _parse_tanks(g: _Grid, hdr: Tuple[int, int], col_max: int, row_limit: int = 10_000) -> TankTable:
    tcol, hrow = hdr
    opens, closes, sales_col, gear = [], [], None, []
    for c in range(tcol + 1, min(col_max, tcol + 14) + 1):
        t = g.t(c, hrow)
        if t in OPEN_DIP:
            opens.append(c)
        elif t in CLOSE_DIP:
            closes.append(c)
        elif t == "SALES" and sales_col is None:
            sales_col = c
        elif "GEAR" in t or t == "DIFFERENCE":
            gear.append(c)
    pairs = []
    for o in opens:
        cands = [c for c in closes if c > o]
        if cands:
            pairs.append((o, cands[0]))
    if not pairs:
        raise LayoutError(f"no dipping columns next to TANK at {coord(tcol, hrow)}")
    tbl = TankTable(hrow, tcol, pairs, sales_col, gear, row_limit=row_limit)

    stop_tokens = ADD_UP | AVAILABLE_CASH | EXPENSES | IMPREST | BANK_LODGMENT | PRODUCTS_RECEIVED
    for r in range(hrow + 1, min(g.max_row + 1, row_limit)):
        # Stop if this row is the POS / ADD UP header row
        if any(g.t(cc, r) in (ADD_UP | AVAILABLE_CASH) for cc in range(max(1, tcol - 3), min(col_max, tcol + 14) + 1)):
            break
        lab = g.v(tcol, r)
        if lab is None or str(lab).strip() == "" or is_formula(lab):
            continue
        n = norm(lab)
        if n in stop_tokens or n.startswith("TOTAL"):
            break
        has_dip = any(
            is_number(g.v(c, r)) or is_formula(g.v(c, r)) for p in pairs for c in p
        )
        has_sales = sales_col is not None and is_formula(g.v(sales_col, r))
        prod = product_key(lab)
        # Accept any tank row with dipping values, sales formula, or a valid product prefix (even if newly added and blank)
        if not (has_dip or has_sales or bool(prod)):
            continue
        tbl.tanks.append(Tank(str(lab).strip(), r, prod, has_sales_formula=has_sales))

    # manifold detection: an opening/closing formula in the *effective*
    # (last) pair that references another tank row, e.g. F29 = D29 + D35.
    # Disambiguate duplicate tank labels (e.g. Sango has two 'AGO 1' rows).
    seen_labels: Dict[str, int] = {}
    for t in tbl.tanks:
        k = compact(t.label)
        seen_labels[k] = seen_labels.get(k, 0) + 1
        if seen_labels[k] > 1:
            t.label = f"{t.label} ({seen_labels[k]})"

    row_to_tank = {t.row: t for t in tbl.tanks}
    eff_open, eff_close = pairs[-1]
    for t in tbl.tanks:
        for c in (eff_open, eff_close):
            f = g.v(c, t.row)
            if not is_formula(f):
                continue
            for _, rr in _REF_RE.findall(f):
                rr = int(rr)
                if rr != t.row and rr in row_to_tank:
                    other = row_to_tank[rr]
                    if other.label not in t.hosts:
                        t.hosts.append(other.label)
                    other.linked_into = t.label
    # single-pair manifold: the host's SALES formula sums two tank rows
    if sales_col is not None and len(pairs) == 1:
        for t in tbl.tanks:
            f = g.v(sales_col, t.row)
            if not is_formula(f):
                continue
            for col_letters, rr in _REF_RE.findall(f):
                rr = int(rr)
                if rr != t.row and rr in row_to_tank and rr > hrow:
                    other = row_to_tank[rr]
                    if other.label not in t.hosts:
                        t.hosts.append(other.label)
                    other.linked_into = t.label
    return tbl


def _parse_labeled(
    g: _Grid,
    anchor: Tuple[int, int],
    stop_names=(),
    stop_row: Optional[int] = None,
    require_product: bool = False,
) -> LabeledValues:
    c, r0 = anchor
    lv = LabeledValues(r0, c, c + 1, stop_row=stop_row)
    limit = min(g.max_row, (stop_row - 1) if stop_row else g.max_row)
    seen_labels: Dict[str, int] = {}
    blank_streak = 0
    for r in range(r0 + 1, limit + 1):
        lab = g.v(c, r)
        if lab is None or str(lab).strip() == "" or is_formula(lab):
            blank_streak += 1
            if not stop_row and blank_streak >= 1:
                break
            if blank_streak >= 3:
                break
            continue
        n = norm(lab)
        if n in stop_names or n.startswith("TOTAL") or n in (ADD_UP | AVAILABLE_CASH | EXPENSES | IMPREST | BANK_LODGMENT):
            break
        if require_product and not product_key(lab) and not is_number(g.v(c + 1, r)):
            break
        blank_streak = 0
        raw_label = str(lab).strip()
        k = compact(raw_label)
        seen_labels[k] = seen_labels.get(k, 0) + 1
        final_label = f"{raw_label} ({seen_labels[k]})" if seen_labels[k] > 1 else raw_label
        lv.rows[final_label] = r
    return lv


def _parse_list(g: _Grid, anchor: Tuple[int, int], col_max: int) -> ListBlock:
    c, r0 = anchor
    amount_col = None
    for cc in range(c + 1, min(c + 4, col_max) + 1):
        if g.t(cc, r0) == "AMOUNT":
            amount_col = cc
            break
    if amount_col is None:
        amount_col = c + 1
    total = None
    for r in range(r0 + 1, min(r0 + 40, g.max_row) + 1):
        f = g.v(amount_col, r)
        if is_formula(f) and "SUM" in f.upper():
            total = r
            break
        if any(g.t(cc, r) == "TOTAL" for cc in range(c, amount_col + 1)):
            total = r
            break
    last = (total - 1) if total else r0 + 1
    return ListBlock(r0, c, amount_col, r0 + 1, last, total)


def _parse_pos(g: _Grid, addup: Tuple[int, int], avail_cols: List[int], col_min: int = 1) -> PosTable:
    from openpyxl.utils import column_index_from_string

    ac, hrow = addup
    avail = next((c for c in avail_cols if c == ac + 1), None)
    if avail is None:
        hits = [c for (c, r) in g.find(AVAILABLE_CASH, row_min=hrow, row_max=hrow) if c > ac]
        avail = hits[0] if hits else None
    vrow = hrow + 1

    # Also inspect the AVAILABLE CASH formula on vrow to see every POS column subtracted
    formula_cols: List[int] = []
    if avail is not None:
        f = g.v(avail, vrow)
        if is_formula(f):
            for col_str, r_str in _REF_RE.findall(f):
                if int(r_str) == vrow:
                    cc = column_index_from_string(col_str)
                    if col_min <= cc < ac:
                        formula_cols.append(cc)

    scan_min = min(formula_cols) if formula_cols else col_min
    channels: List[Tuple[str, int]] = []
    all_cols: List[int] = []
    c = ac - 1
    blank_streak = 0
    while c >= max(1, scan_min):
        h = g.v(c, hrow)
        v = g.v(c, vrow)
        if norm(h) in AVAILABLE_CASH or norm(h) in ADD_UP:
            break
        if h is None and v is None and c not in formula_cols:
            blank_streak += 1
            if blank_streak >= 2 and c < scan_min:
                break
            if not formula_cols and blank_streak >= 1:
                break
            c -= 1
            continue
        blank_streak = 0
        all_cols.append(c)
        if h is not None and str(h).strip() not in ("", "0"):
            channels.append((str(h).strip(), c))
        c -= 1
    channels.reverse()
    all_cols.reverse()
    # Deduplicate identical channel names within the same half (e.g. two 'cash/transfer')
    seen: Dict[str, int] = {}
    dedup_channels: List[Tuple[str, int]] = []
    for label, col in channels:
        k = norm(label)
        seen[k] = seen.get(k, 0) + 1
        if seen[k] > 1:
            label = f"{label} ({seen[k]})"
        dedup_channels.append((label, col))
    return PosTable(hrow, vrow, dedup_channels, ac, avail, all_cols)


def _date_cells(ws: Worksheet, col_min: int, col_max: int) -> List[str]:
    out = []
    for r in (1, 2):
        for c in range(col_min, min(col_max, ws.max_column) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (_dt.datetime, _dt.date)):
                out.append(coord(c, r))
    return out


def parse_sheet(ws: Worksheet) -> SheetLayout:
    g = _Grid(ws)
    opens = g.find(OPENING_METER, row_max=8)
    if not opens:
        # Typo fallback (e.g. Ogbomoso 2 has 'CLOSING METER' in the Opening Meter header cell
        # followed by 4 more 'CLOSING METER' headers).
        for (c, r) in g.find(CLOSING_METER, row_max=8):
            if g.t(c - 1, r) not in CLOSING_METER and g.t(c + 1, r) in CLOSING_METER and g.t(c + 2, r) in CLOSING_METER:
                opens.append((c, r))
    if not opens:
        raise LayoutError(f"{ws.title}: no 'OPENING METER' anchor found")
    opens.sort()
    left_open = opens[0]
    right_open = next((o for o in opens[1:] if o[0] > left_open[0] + 8), None)
    # Columns are NOT a clean split (e.g. Lekki 3's right TANK header sits at
    # AC while the left IMPREST block sits at AA).  So every anchor kind is
    # paired instead: with two hits, left = left-most, right = right-most; a
    # single hit belongs to the half it is closest to.
    split = (right_open[0] - 3) if right_open else g.max_col + 1

    def pick(hits, side):
        hits = sorted(hits)
        if not hits:
            return None
        if right_open is None:
            return hits[0] if side == "LEFT" else None
        if len(hits) >= 2:
            return hits[0] if side == "LEFT" else hits[-1]
        h = hits[0]
        is_left = h[0] < split
        return h if (side == "LEFT") == is_left else None

    tank_hdrs = [
        t for t in g.find(TANK)
        if any(g.t(c, t[1]) in OPEN_DIP for c in range(t[0] + 1, t[0] + 4))
    ]
    if not tank_hdrs:
        # Fallback when 'TANK' cell is blank (e.g. Igbesa ags) but 'OPENNG DIPPG' + 'CLOSNG DIPPG' exist
        for (c, r) in g.find(OPEN_DIP):
            if g.t(c - 1, r) not in OPEN_DIP and g.t(c - 1, r) not in CLOSE_DIP and g.t(c + 1, r) in CLOSE_DIP:
                tank_hdrs.append((c - 1, r))
    addups = g.find(ADD_UP)
    avails = [c for c, _ in g.find(AVAILABLE_CASH)]

    def build(side, op, cmin, cmax):
        warnings = []
        tank_hdr = pick([t for t in tank_hdrs if t[1] > op[1]], side)
        row_stop = tank_hdr[1] if tank_hdr else g.max_row
        sales = _parse_sales(g, op[0], op[1], cmax, row_stop)
        a = pick(addups, side)
        tanks = None
        if tank_hdr:
            try:
                limit = a[1] if (a and a[1] > tank_hdr[1]) else 10_000
                tanks = _parse_tanks(g, tank_hdr, cmax, limit)
            except LayoutError as e:
                warnings.append(str(e))
        else:
            warnings.append("TANK table not found")

        rec_a = pick(g.find(PRODUCTS_RECEIVED), side)
        exp_a = pick(g.find(EXPENSES), side)
        rec_stop_row = exp_a[1] if (rec_a and exp_a and exp_a[1] > rec_a[1]) else None
        received = (
            _parse_labeled(g, rec_a, stop_names=EXPENSES, stop_row=rec_stop_row, require_product=True)
            if rec_a
            else None
        )
        expenses = _parse_list(g, exp_a, g.max_col) if exp_a else None
        imp_a = pick(g.find(IMPREST), side)
        imprest = _parse_list(g, imp_a, g.max_col) if imp_a else None
        bank_a = pick(g.find(BANK_LODGMENT), side)
        bank = _parse_labeled(g, bank_a) if bank_a else None

        # Link each Tank to its exact PRODUCTS RECEIVED row when referenced in the Tank SALES formula
        if tanks and tanks.sales_col and received:
            rec_col_let = get_column_letter(received.value_col)
            for tk in tanks.tanks:
                sf = g.v(tanks.sales_col, tk.row)
                if is_formula(sf):
                    for c_let, r_str in _REF_RE.findall(sf):
                        if c_let == rec_col_let:
                            tk.received_row = int(r_str)
                            break

        pos = None
        if a:
            pos = _parse_pos(g, a, avails, col_min=cmin)
        else:
            warnings.append("ADD UP / POS table not found")
        if received is None:
            warnings.append("PRODUCTS RECEIVED not found")
        return HalfLayout(side, cmin, cmax, _date_cells(ws, cmin, cmax), sales, tanks,
                          received, expenses, imprest, pos, bank, warnings)

    left = build("LEFT", left_open, 1, (right_open[0] - 1) if right_open else g.max_col)
    right = build("RIGHT", right_open, split, g.max_col) if right_open else None
    if right is not None and len(right.pos.channels if right.pos else []) == 0:
        right.warnings.append("RIGHT half has no POS channels")
    return SheetLayout(ws.title, left, right)
