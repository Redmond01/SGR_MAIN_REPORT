"""Core rollover, dynamic tank adaptation, mirroring, and write-back engine.

Given any of the ~60 station workbooks:
1. ``inspect_workbook`` parses the latest sheet and builds a :class:`StationState`
   holding the rolled-over opening figures, active prices, tank manifold links,
   and POS channels so the UI can render a clean data-entry form.
2. ``process_workbook`` duplicates the latest sheet as ``Sheet{N+1}``, performs
   the full daily rollover + reset, applies the user's :class:`StationInput`
   (including mid-day price changes and dynamic tank maintenance/manifolding),
   mirrors Tank/Delivery/POS/Price data onto the Right (Inner Meter) half, and
   returns the updated workbook + live evaluated summary metrics.
"""
from __future__ import annotations

import datetime as _dt
import gc
import io
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .evaluator import SheetEvaluator
from .parser import (
    HalfLayout,
    Pump,
    SheetLayout,
    Tank,
    TankTable,
    compact,
    coord,
    is_formula,
    is_number,
    norm,
    parse_sheet,
)

__all__ = [
    "PumpRowState",
    "TankRowState",
    "PosChannelState",
    "ExpenseEntry",
    "StationState",
    "StationSummary",
    "next_sheet_name",
    "inspect_workbook",
    "rollover_sheet",
    "apply_station_state",
    "process_workbook",
]


# ---------------------------------------------------------------- data models


@dataclass
class PumpRowState:
    label: str
    product: str
    row: int
    opening_meter: float
    closing_meter_1: float
    closing_meter_2: float
    price_1: float
    price_2: float = 0.0
    inner_opening_meter: float = 0.0
    inner_closing_meter_1: float = 0.0
    inner_closing_meter_2: float = 0.0
    closing_meter_3: float = 0.0
    closing_meter_4: float = 0.0
    price_3: float = 0.0
    price_4: float = 0.0
    inner_closing_meter_3: float = 0.0
    inner_closing_meter_4: float = 0.0

    def get_closing(self, tier: int) -> float:
        return getattr(self, f"closing_meter_{tier}")

    def set_closing(self, tier: int, val: float):
        setattr(self, f"closing_meter_{tier}", val)

    def get_price(self, tier: int) -> float:
        return getattr(self, f"price_{tier}")

    def set_price(self, tier: int, val: float):
        setattr(self, f"price_{tier}", val)

    def get_inner_closing(self, tier: int) -> float:
        return getattr(self, f"inner_closing_meter_{tier}")

    def set_inner_closing(self, tier: int, val: float):
        setattr(self, f"inner_closing_meter_{tier}", val)

    def compute_tiers(self, tiers: int) -> dict:
        """Compute per-tier Closing, Sales, Price, and Amount across ``1..tiers``."""
        tiers = max(1, min(4, int(tiers)))
        prev_c = self.opening_meter
        prev_ic = self.inner_opening_meter
        prev_p = self.price_1
        tier_rows = []
        tot_l = 0.0
        tot_amt = 0.0
        itot_l = 0.0
        itot_amt = 0.0
        for t in range(1, tiers + 1):
            raw_c = self.get_closing(t)
            c_val = raw_c if raw_c >= prev_c else prev_c
            raw_ic = self.get_inner_closing(t)
            ic_val = raw_ic if raw_ic >= prev_ic else prev_ic
            raw_p = self.get_price(t)
            p_val = raw_p if raw_p > 0 else prev_p
            if p_val > 0:
                prev_p = p_val
            s_val = round(max(0.0, c_val - prev_c), 2)
            is_val = round(max(0.0, ic_val - prev_ic), 2)
            amt_val = round(s_val * p_val, 2)
            iamt_val = round(is_val * p_val, 2)
            tot_l += s_val
            tot_amt += s_val * p_val
            itot_l += is_val
            itot_amt += is_val * p_val
            tier_rows.append({
                "tier": t,
                "closing": round(c_val, 2),
                "sales": s_val,
                "price": p_val,
                "amount": amt_val,
                "inner_closing": round(ic_val, 2),
                "inner_sales": is_val,
                "inner_amount": iamt_val,
            })
            prev_c = c_val
            prev_ic = ic_val
        return {
            "tiers": tier_rows,
            "total_sales_ltr": round(tot_l, 2),
            "total_amount": round(tot_amt, 2),
            "inner_total_ltr": round(itot_l, 2),
            "inner_total_amount": round(itot_amt, 2),
        }


@dataclass
class TankRowState:
    label: str
    product: str
    row: int
    opening_dip: float
    closing_dip: float
    product_received: float = 0.0
    under_maintenance: bool = False
    linked_into: Optional[str] = None  # label of host tank if manifolded


@dataclass
class PosChannelState:
    label: str
    col: int
    amount: float = 0.0


@dataclass
class ExpenseEntry:
    description: str = ""
    amount: float = 0.0


@dataclass
class StationSummary:
    station_name: str
    source_sheet: str
    target_sheet: str
    total_sales_amount: float
    inner_total_sales_amount: float
    product_litres: Dict[str, float]
    total_expenses: float
    total_pos: float
    add_up: float
    available_cash: float
    inner_available_cash: float
    warnings: List[str] = field(default_factory=list)


@dataclass
class StationState:
    file_name: str
    station_name: str
    source_sheet: str
    target_sheet: str
    report_date: _dt.date
    price_change_day: bool
    pumps: List[PumpRowState]
    tanks: List[TankRowState]
    pos_channels: List[PosChannelState]
    add_up: float
    inner_add_up: float
    expenses: List[ExpenseEntry] = field(default_factory=list)
    imprest: List[ExpenseEntry] = field(default_factory=list)
    bank_name: str = ""
    bank_amount_deposited: float = 0.0
    bank_payment_status: str = ""
    mirror_pos_to_inner: bool = True
    price_tiers: int = 1  # 1 = standard day, 2 = 1 price change, 3 = 2 price changes, 4 = 3 price changes
    is_saved: bool = False
    saved_at: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
    source_sheet_count: int = 0
    validation_errors: List[str] = field(default_factory=list)
    validation_warnings: List[str] = field(default_factory=list)

    def __post_init__(self):
        if self.price_change_day and self.price_tiers < 2:
            self.price_tiers = 2
        elif not self.price_change_day and self.price_tiers > 1:
            self.price_change_day = True

    def set_price_tiers(self, tiers: int):
        tiers = max(1, min(4, int(tiers)))
        self.price_tiers = tiers
        self.price_change_day = tiers > 1
        for p in self.pumps:
            prev_c = p.opening_meter
            prev_ic = p.inner_opening_meter
            prev_p = p.price_1
            for t in range(1, 5):
                if t <= tiers:
                    cur_c = p.get_closing(t)
                    if cur_c < prev_c:
                        cur_c = prev_c
                        p.set_closing(t, cur_c)
                    prev_c = cur_c

                    cur_ic = p.get_inner_closing(t)
                    if cur_ic < prev_ic:
                        cur_ic = prev_ic
                        p.set_inner_closing(t, cur_ic)
                    prev_ic = cur_ic

                    cur_p = p.get_price(t)
                    if cur_p <= 0 and prev_p > 0:
                        cur_p = prev_p
                        p.set_price(t, cur_p)
                    elif cur_p > 0:
                        prev_p = cur_p


# ---------------------------------------------------------------- helpers


_SHEET_NUM_RE = re.compile(r"^(.*?)(\d+)$")


def next_sheet_name(wb: Workbook) -> str:
    """Return ``Sheet{N+1}`` based on the last worksheet's name."""
    last = wb.worksheets[-1].title
    m = _SHEET_NUM_RE.match(last.strip())
    if m:
        prefix, num = m.group(1), int(m.group(2))
        cand = f"{prefix}{num + 1}"
    else:
        cand = f"{last}_1"
    existing = set(wb.sheetnames)
    i = 2
    base = cand
    while cand in existing:
        cand = f"{base}_{i}"
        i += 1
    return cand


def clean_station_name(file_name: str) -> str:
    name = re.sub(r"\.xlsx$", "", file_name, flags=re.I)
    name = re.sub(
        r"^(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+",
        "",
        name,
        flags=re.I,
    )
    name = re.sub(r"\s*\(\d+\)\s*", " ", name)
    return re.sub(r"\s+", " ", name).strip()


def _num(v, default: float = 0.0) -> float:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return default


def _last_closing_value(ws: Worksheet, evalr: SheetEvaluator, sales_tbl, row: int) -> float:
    """Return yesterday's final closing meter reading for ``row``."""
    vals: List[float] = []
    for c in sales_tbl.closing_cols:
        raw = ws.cell(row=row, column=c).value
        if is_number(raw) and float(raw) != 0.0:
            vals.append(float(raw))
        elif is_formula(raw):
            v = evalr.try_value(c, row)
            if v is not None and v != 0.0:
                vals.append(v)
    if vals:
        return vals[-1]
    return _num(ws.cell(row=row, column=sales_tbl.open_col).value, 0.0)


def _active_price_value(ws: Worksheet, sales_tbl, row: int) -> float:
    """Return yesterday's final active price for ``row`` across up to 4 tiers."""
    pcols = sales_tbl.price_cols
    ccols = sales_tbl.closing_cols
    n = min(len(pcols), len(ccols))
    for i in range(n - 1, 0, -1):
        pi = ws.cell(row=row, column=pcols[i]).value
        ci = ws.cell(row=row, column=ccols[i]).value
        if is_number(pi) and float(pi) > 0 and is_number(ci) and float(ci) > 0:
            return float(pi)
    if pcols:
        p1 = ws.cell(row=row, column=pcols[0]).value
        if is_number(p1) and float(p1) > 0:
            return float(p1)
    return 0.0


def _physical_dip_pair(ws: Worksheet, tank_tbl: TankTable, tk: Tank) -> Tuple[int, int]:
    """Return the ``(open_col, close_col)`` pair holding physical dippings for ``tk``."""
    pairs = tank_tbl.dip_pairs
    if len(pairs) == 1:
        return pairs[0]
    eff_o, eff_c = pairs[-1]
    raw_eff_o = ws.cell(row=tk.row, column=eff_o).value
    raw_eff_c = ws.cell(row=tk.row, column=eff_c).value
    if is_formula(raw_eff_o) or is_formula(raw_eff_c) or tk.linked_into or tk.hosts:
        return pairs[0]
    raw_p0_o = ws.cell(row=tk.row, column=pairs[0][0]).value
    raw_p0_c = ws.cell(row=tk.row, column=pairs[0][1]).value
    if (is_number(raw_p0_o) or is_number(raw_p0_c)) and raw_eff_o is None and raw_eff_c is None:
        return pairs[0]
    return pairs[-1]


def _match_received_row(
    half: HalfLayout, tank_label: str, tank_idx: int, tk: Optional[Tank] = None
) -> Optional[int]:
    """Find the row in ``PRODUCTS RECEIVED`` corresponding to ``tank_label``."""
    if not half.received or not half.received.rows:
        return None
    if tk is not None and tk.received_row is not None:
        return tk.received_row
    target = compact(tank_label)
    for k, r in half.received.rows.items():
        if compact(k) == target:
            return r
    # Match single-tank product family (e.g. 'AGO 1' <-> 'AGO', 'LPG 1' <-> 'LPG')
    for k, r in half.received.rows.items():
        ck = compact(k)
        if ck and (target.startswith(ck) or ck.startswith(target)):
            return r
    # Fallback by index if within range
    items = list(half.received.rows.values())
    if tank_idx < len(items):
        return items[tank_idx]
    return None


def _ensure_received_row(
    ws: Worksheet, half: HalfLayout, tank_label: str, tank_idx: int, tk: Optional[Tank] = None
) -> Optional[int]:
    """Return existing or newly allocated row in ``PRODUCTS RECEIVED`` for ``tank_label``."""
    rr = _match_received_row(half, tank_label, tank_idx, tk=tk)
    if rr is not None or not half.received:
        return rr
    rec = half.received
    last_r = max(rec.rows.values()) if rec.rows else rec.anchor_row
    cand_r = last_r + 1
    if rec.stop_row and cand_r >= rec.stop_row:
        return None
    if ws.cell(row=cand_r, column=rec.label_col).value is None:
        ws.cell(row=cand_r, column=rec.label_col).value = tank_label
        rec.rows[tank_label] = cand_r
        if tk is not None:
            tk.received_row = cand_r
        return cand_r
    return None


# ---------------------------------------------------------------- inspection


def inspect_workbook(
    source: Union[bytes, str, io.BytesIO],
    file_name: str = "station.xlsx",
    report_date: Optional[_dt.date] = None,
) -> StationState:
    """Inspect a station workbook and build the rolled-over :class:`StationState`."""
    if isinstance(source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(source), data_only=False)
    elif isinstance(source, io.BytesIO):
        wb = openpyxl.load_workbook(io.BytesIO(source.getvalue()), data_only=False)
    else:
        wb = openpyxl.load_workbook(source, data_only=False)

    ws = wb.worksheets[-1]
    layout = parse_sheet(ws)
    evalr = SheetEvaluator(ws)

    # Infer default report_date = yesterday's sheet date + 1 day
    inferred_date = report_date
    if inferred_date is None:
        for dc in layout.left.date_cells:
            dv = ws[dc].value
            if isinstance(dv, _dt.datetime):
                inferred_date = dv.date() + _dt.timedelta(days=1)
                break
            if isinstance(dv, _dt.date):
                inferred_date = dv + _dt.timedelta(days=1)
                break
    if inferred_date is None:
        inferred_date = _dt.date.today()

    # Build PumpRowState list (rolling over yesterday's closing -> today's opening)
    pumps: List[PumpRowState] = []
    lsales = layout.left.sales
    rsales = layout.right.sales if layout.right else None

    for idx, p in enumerate(lsales.pumps):
        open_m = _last_closing_value(ws, evalr, lsales, p.row)
        price_1 = _active_price_value(ws, lsales, p.row)
        inner_open = 0.0
        if rsales:
            rp = (rsales.pumps[idx] if idx < len(rsales.pumps) else None) or rsales.pump(p.label)
            if rp:
                inner_open = _last_closing_value(ws, evalr, rsales, rp.row)
        pumps.append(
            PumpRowState(
                label=p.label,
                product=p.product,
                row=p.row,
                opening_meter=round(open_m, 2),
                closing_meter_1=round(open_m, 2),
                closing_meter_2=round(open_m, 2),
                closing_meter_3=round(open_m, 2),
                closing_meter_4=round(open_m, 2),
                price_1=price_1,
                price_2=0.0,
                price_3=0.0,
                price_4=0.0,
                inner_opening_meter=round(inner_open, 2),
                inner_closing_meter_1=round(inner_open, 2),
                inner_closing_meter_2=round(inner_open, 2),
                inner_closing_meter_3=round(inner_open, 2),
                inner_closing_meter_4=round(inner_open, 2),
            )
        )

    # Build TankRowState list (rolling over yesterday's closing dip -> today's opening dip)
    # Also cross-check PRODUCTS RECEIVED so if a station increases in tanks in either section, it is picked up.
    tanks: List[TankRowState] = []
    if layout.left.tanks:
        ltanks = layout.left.tanks
        for idx, tk in enumerate(ltanks.tanks):
            o_col, c_col = _physical_dip_pair(ws, ltanks, tk)
            prev_close = ws.cell(row=tk.row, column=c_col).value
            if is_formula(prev_close):
                dip_val = evalr.try_value(c_col, tk.row) or 0.0
            elif is_number(prev_close):
                dip_val = float(prev_close)
            else:
                prev_open = ws.cell(row=tk.row, column=o_col).value
                dip_val = float(prev_open) if is_number(prev_open) else 0.0
            tanks.append(
                TankRowState(
                    label=tk.label,
                    product=tk.product,
                    row=tk.row,
                    opening_dip=round(dip_val, 2),
                    closing_dip=round(dip_val, 2),
                    product_received=0.0,
                    under_maintenance=False,
                    linked_into=tk.linked_into,
                )
            )
    if layout.left.received and layout.left.received.rows:
        existing_compact = {compact(t.label) for t in tanks}
        for rec_label, rec_row in layout.left.received.rows.items():
            ck = compact(rec_label)
            if not ck or ck in existing_compact:
                continue
            # Check if already covered by a prefix match (e.g. 'AGO' vs 'AGO 1' or 'LPG' vs 'LPG 1')
            if any(ek.startswith(ck) or ck.startswith(ek) for ek in existing_compact):
                continue
            from .parser import product_key
            prod = product_key(rec_label)
            if prod:
                next_row = (tanks[-1].row + 1) if tanks else rec_row
                tanks.append(
                    TankRowState(
                        label=rec_label,
                        product=prod,
                        row=next_row,
                        opening_dip=0.0,
                        closing_dip=0.0,
                        product_received=0.0,
                        under_maintenance=False,
                        linked_into=None,
                    )
                )
                existing_compact.add(ck)

    # Build PosChannelState & rolled-over ADD UP
    pos_channels: List[PosChannelState] = []
    add_up = 0.0
    if layout.left.pos:
        lpos = layout.left.pos
        for ch_label, col in lpos.channels:
            pos_channels.append(PosChannelState(label=ch_label, col=col, amount=0.0))
        if lpos.available_col:
            v = evalr.try_value(lpos.available_col, lpos.value_row)
            if v is not None:
                add_up = round(v, 2)
        else:
            add_up = round(_num(ws.cell(row=lpos.value_row, column=lpos.add_up_col).value), 2)

    inner_add_up = 0.0
    if layout.right and layout.right.pos:
        rpos = layout.right.pos
        # Check if right-half POS was actively mirrored on the previous sheet
        if rpos.available_col:
            rv = evalr.try_value(rpos.available_col, rpos.value_row)
            inner_add_up = round(rv, 2) if rv is not None else 0.0

    # Bank lodgment info
    bank_name = ""
    if layout.left.bank:
        for k, r in layout.left.bank.rows.items():
            if "BANK" in norm(k):
                v = ws.cell(row=r, column=layout.left.bank.value_col).value
                if v is not None:
                    bank_name = str(v).strip()

    warnings = list(layout.left.warnings)
    if layout.right:
        warnings.extend(f"Right half: {w}" for w in layout.right.warnings)

    source_title = ws.title
    target_title = next_sheet_name(wb)
    sheet_count = len(wb.worksheets)
    try:
        wb.close()
    except Exception:
        pass
    del wb, ws, evalr
    gc.collect()

    return StationState(
        file_name=file_name,
        station_name=clean_station_name(file_name),
        source_sheet=source_title,
        target_sheet=target_title,
        report_date=inferred_date,
        price_change_day=False,
        pumps=pumps,
        tanks=tanks,
        pos_channels=pos_channels,
        add_up=add_up,
        inner_add_up=inner_add_up,
        expenses=[],
        imprest=[],
        bank_name=bank_name,
        warnings=warnings,
        source_sheet_count=sheet_count,
    )


# ---------------------------------------------------------------- rollover & write-back


def _reset_half_on_rollover(
    ws: Worksheet,
    prev_ws: Worksheet,
    prev_eval: SheetEvaluator,
    half: HalfLayout,
    report_date: _dt.date,
    roll_pos_addup: bool = True,
):
    """Perform pure rollover/reset on one half of the newly duplicated sheet."""
    dt_val = _dt.datetime(report_date.year, report_date.month, report_date.day)
    for dc in half.date_cells:
        ws[dc].value = dt_val

    # 1. Sales table rollover
    s = half.sales
    open_let = get_column_letter(s.open_col)
    for p in s.pumps:
        final_close = _last_closing_value(prev_ws, prev_eval, s, p.row)
        active_price = _active_price_value(prev_ws, s, p.row)
        ws.cell(row=p.row, column=s.open_col).value = final_close
        if len(s.closing_cols) >= 1:
            ws.cell(row=p.row, column=s.closing_cols[0]).value = final_close
            if len(s.sales_cols) >= 1:
                c1_let = get_column_letter(s.closing_cols[0])
                ws.cell(row=p.row, column=s.sales_cols[0]).value = f"={c1_let}{p.row}-{open_let}{p.row}"
        if len(s.closing_cols) >= 2:
            ws.cell(row=p.row, column=s.closing_cols[1]).value = final_close
            if len(s.sales_cols) >= 2:
                c1_let = get_column_letter(s.closing_cols[0])
                c2_let = get_column_letter(s.closing_cols[1])
                ws.cell(row=p.row, column=s.sales_cols[1]).value = f"={c2_let}{p.row}-{c1_let}{p.row}"
        for c in s.closing_cols[2:]:
            ws.cell(row=p.row, column=c).value = None
        # Restore standard-day Sales 3 & Sales 4 formulas (=G{row} and =H{row}-G{row})
        if len(s.sales_cols) >= 3 and len(s.closing_cols) >= 3:
            c3_let = get_column_letter(s.closing_cols[2])
            ws.cell(row=p.row, column=s.sales_cols[2]).value = f"={c3_let}{p.row}"
        if len(s.sales_cols) >= 4 and len(s.closing_cols) >= 4:
            c3_let = get_column_letter(s.closing_cols[2])
            c4_let = get_column_letter(s.closing_cols[3])
            ws.cell(row=p.row, column=s.sales_cols[3]).value = f"={c4_let}{p.row}-{c3_let}{p.row}"
        if len(s.price_cols) >= 1 and active_price > 0:
            ws.cell(row=p.row, column=s.price_cols[0]).value = active_price
        for c in s.price_cols[1:]:
            if not is_formula(ws.cell(row=p.row, column=c).value):
                ws.cell(row=p.row, column=c).value = None
        if s.amount_col and s.sales_cols and s.price_cols:
            n_pairs = min(len(s.sales_cols), len(s.price_cols))
            terms = [
                f"({get_column_letter(s.sales_cols[i])}{p.row}*{get_column_letter(s.price_cols[i])}{p.row})"
                for i in range(n_pairs)
            ]
            ws.cell(row=p.row, column=s.amount_col).value = "=" + "+".join(terms)

    # 2. Tank table rollover (preserve formulas, roll closing -> opening)
    if half.tanks:
        t = half.tanks
        for tk in t.tanks:
            o_col, c_col = _physical_dip_pair(prev_ws, t, tk)
            prev_c = prev_ws.cell(row=tk.row, column=c_col).value
            if is_number(prev_c):
                ws.cell(row=tk.row, column=o_col).value = prev_c
                ws.cell(row=tk.row, column=c_col).value = prev_c

    # 3. Products Received reset
    if half.received:
        for r in half.received.rows.values():
            c = ws.cell(row=r, column=half.received.value_col)
            if not is_formula(c.value):
                c.value = None

    # 4. Expenses & Imprest reset (keep formula rows untouched)
    for blk in (half.expenses, half.imprest):
        if blk is None:
            continue
        for r in range(blk.first_row, blk.last_row + 1):
            dc = ws.cell(row=r, column=blk.desc_col)
            ac = ws.cell(row=r, column=blk.amount_col)
            if not is_formula(dc.value) and not is_formula(ac.value):
                dc.value = None
                ac.value = None

    # 5. POS & ADD UP rollover
    if half.pos and roll_pos_addup:
        pos = half.pos
        if pos.available_col is not None:
            avail_val = prev_eval.try_value(pos.available_col, pos.value_row)
            if avail_val is not None:
                ws.cell(row=pos.value_row, column=pos.add_up_col).value = round(avail_val, 2)
        for col in pos.all_value_cols:
            cell = ws.cell(row=pos.value_row, column=col)
            if not is_formula(cell.value):
                cell.value = 0


def _apply_tank_links(ws: Worksheet, half: HalfLayout, tank_states: List[TankRowState]):
    """Adapt tank formulas dynamically when manifold links change."""
    if not half.tanks:
        return
    t = half.tanks
    state_by_label = {compact(ts.label): ts for ts in tank_states}
    tank_by_label = {compact(tk.label): tk for tk in t.tanks}

    # Check if any tank's linked_into differs from the sheet's existing links
    changed = False
    for tk in t.tanks:
        ts = state_by_label.get(compact(tk.label))
        if ts is None:
            continue
        cur_link = compact(tk.linked_into) if tk.linked_into else ""
        new_link = compact(ts.linked_into) if ts.linked_into else ""
        if cur_link != new_link:
            changed = True
            break
    if not changed or len(t.dip_pairs) < 2:
        return

    # Rebuild 2-pair manifold structure according to new links
    (o1, c1), (o2, c2) = t.dip_pairs[0], t.dip_pairs[-1]
    hosts_map: Dict[str, List[Tank]] = {compact(tk.label): [] for tk in t.tanks}
    for ts in tank_states:
        if ts.linked_into:
            hk = compact(ts.linked_into)
            tk = tank_by_label.get(compact(ts.label))
            if hk in hosts_map and tk is not None and hk != compact(ts.label):
                hosts_map[hk].append(tk)

    for idx, tk in enumerate(t.tanks):
        key = compact(tk.label)
        ts = state_by_label.get(key)
        children = hosts_map.get(key, [])
        is_child = bool(ts and ts.linked_into and compact(ts.linked_into) in tank_by_label)
        if is_child:
            # Physical dips live in pair 0; pair 1 & sales are cleared
            ws.cell(row=tk.row, column=o2).value = None
            ws.cell(row=tk.row, column=c2).value = None
            if t.sales_col:
                ws.cell(row=tk.row, column=t.sales_col).value = None
            tk.linked_into = ts.linked_into
            tk.hosts = []
        elif children:
            # Host tank: physical dips in pair 0, sum formula in pair 1
            rows = sorted([tk.row] + [ch.row for ch in children])
            ws.cell(row=tk.row, column=o2).value = "=" + "+".join(f"{get_column_letter(o1)}{r}" for r in rows)
            ws.cell(row=tk.row, column=c2).value = "=" + "+".join(f"{get_column_letter(c1)}{r}" for r in rows)
            if t.sales_col and half.received:
                rec_col = get_column_letter(half.received.value_col)
                rec_refs = []
                for r_tk in [tk] + children:
                    r_idx = t.tanks.index(r_tk)
                    rr = _match_received_row(half, r_tk.label, r_idx)
                    if rr:
                        rec_refs.append(f"{rec_col}{rr}")
                rec_part = ("+" + "+".join(rec_refs)) if rec_refs else ""
                ws.cell(row=tk.row, column=t.sales_col).value = (
                    f"={get_column_letter(o2)}{tk.row}{rec_part}-{get_column_letter(c2)}{tk.row}"
                )
            tk.linked_into = None
            tk.hosts = [ch.label for ch in children]
        else:
            # Standalone tank: physical dips in pair 1, pair 0 cleared
            ws.cell(row=tk.row, column=o1).value = None
            ws.cell(row=tk.row, column=c1).value = None
            if t.sales_col and not is_formula(ws.cell(row=tk.row, column=t.sales_col).value):
                rec_part = ""
                if half.received:
                    rr = _match_received_row(half, tk.label, idx)
                    if rr:
                        rec_part = f"+{get_column_letter(half.received.value_col)}{rr}"
                ws.cell(row=tk.row, column=t.sales_col).value = (
                    f"={get_column_letter(o2)}{tk.row}{rec_part}-{get_column_letter(c2)}{tk.row}"
                )
            tk.linked_into = None
            tk.hosts = []


def apply_station_state(ws: Worksheet, layout: SheetLayout, state: StationState):
    """Write all user inputs from ``state`` into rolled-over worksheet ``ws``."""
    dt_val = _dt.datetime(state.report_date.year, state.report_date.month, state.report_date.day)
    for h in layout.halves:
        for dc in h.date_cells:
            ws[dc].value = dt_val
    # Also set B1 or A1 if date_cells was empty
    if not layout.left.date_cells:
        ws["B1"].value = dt_val

    lsales = layout.left.sales
    rsales = layout.right.sales if layout.right else None
    tiers = max(1, min(4, int(state.price_tiers if state.price_tiers > 1 else (2 if state.price_change_day else 1))))

    def _write_pump_tier_row(sales_tbl, row_num: int, open_val: float, closings: List[float], prices: List[float]):
        ws.cell(row=row_num, column=sales_tbl.open_col).value = open_val
        ccols = sales_tbl.closing_cols
        scols = sales_tbl.sales_cols
        pcols = sales_tbl.price_cols
        open_let = get_column_letter(sales_tbl.open_col)

        # Ensure progressive monotonicity across active tiers & non-zero active prices
        active_c: List[float] = []
        active_p: List[float] = []
        prev_c = open_val
        last_p = prices[0] if (prices and prices[0] > 0) else _active_price_value(ws, sales_tbl, row_num)
        for t in range(tiers):
            raw_c = closings[t] if (t < len(closings) and closings[t] >= prev_c) else prev_c
            active_c.append(raw_c)
            prev_c = raw_c

            raw_p = prices[t] if (t < len(prices) and prices[t] > 0) else last_p
            if raw_p > 0:
                last_p = raw_p
            active_p.append(raw_p)

        # Handle single-column layout (e.g. Interchange) when tiers > 1
        if len(ccols) == 1 and tiers > 1:
            ws.cell(row=row_num, column=ccols[0]).value = active_c[-1]
            if scols:
                c1_let = get_column_letter(ccols[0])
                ws.cell(row=row_num, column=scols[0]).value = f"={c1_let}{row_num}-{open_let}{row_num}"
            if pcols:
                ws.cell(row=row_num, column=pcols[0]).value = active_p[-1] if active_p[-1] > 0 else None
            if sales_tbl.amount_col:
                tier_terms = []
                pc = open_val
                for t in range(tiers):
                    s_t = round(max(0.0, active_c[t] - pc), 2)
                    tier_terms.append(f"({s_t}*{active_p[t]})")
                    pc = active_c[t]
                ws.cell(row=row_num, column=sales_tbl.amount_col).value = "=" + "+".join(tier_terms)
            return

        # Write Closing Meter columns and Sales subtraction formulas
        for i, ccol in enumerate(ccols):
            if i < tiers:
                ws.cell(row=row_num, column=ccol).value = active_c[i]
                if i < len(scols):
                    c_let = get_column_letter(ccol)
                    prev_let = open_let if i == 0 else get_column_letter(ccols[i - 1])
                    ws.cell(row=row_num, column=scols[i]).value = f"={c_let}{row_num}-{prev_let}{row_num}"
            elif i == 1 and tiers == 1:
                # Standard day: Closing 2 == Closing 1 so Sales 2 = C2 - C1 = 0
                ws.cell(row=row_num, column=ccol).value = active_c[0]
                if i < len(scols):
                    c1_let = get_column_letter(ccols[0])
                    c2_let = get_column_letter(ccol)
                    ws.cell(row=row_num, column=scols[i]).value = f"={c2_let}{row_num}-{c1_let}{row_num}"
            elif i == 3 and tiers == 3:
                # 3-tier day: Closing 4 == Closing 3 so Sales 4 = C4 - C3 = 0
                ws.cell(row=row_num, column=ccol).value = active_c[2]
                if i < len(scols):
                    c3_let = get_column_letter(ccols[2])
                    c4_let = get_column_letter(ccol)
                    ws.cell(row=row_num, column=scols[i]).value = f"={c4_let}{row_num}-{c3_let}{row_num}"
            else:
                ws.cell(row=row_num, column=ccol).value = None
                if i == 2 and i < len(scols):
                    c3_let = get_column_letter(ccol)
                    ws.cell(row=row_num, column=scols[i]).value = f"={c3_let}{row_num}"
                elif i == 3 and i < len(scols) and len(ccols) >= 4:
                    c3_let = get_column_letter(ccols[2])
                    c4_let = get_column_letter(ccol)
                    ws.cell(row=row_num, column=scols[i]).value = f"={c4_let}{row_num}-{c3_let}{row_num}"

        # Write Price columns
        for i, pcol in enumerate(pcols):
            if i < tiers:
                pval = active_p[i] if i < len(active_p) else last_p
                ws.cell(row=row_num, column=pcol).value = pval if pval > 0 else None
            else:
                if not is_formula(ws.cell(row=row_num, column=pcol).value):
                    ws.cell(row=row_num, column=pcol).value = None

        # Always write/repair the full multi-tier AMOUNT formula: =(S1*P1)+(S2*P2)+(S3*P3)+(S4*P4)
        if sales_tbl.amount_col and scols and pcols:
            n_pairs = min(len(scols), len(pcols))
            terms = [
                f"({get_column_letter(scols[i])}{row_num}*{get_column_letter(pcols[i])}{row_num})"
                for i in range(n_pairs)
            ]
            ws.cell(row=row_num, column=sales_tbl.amount_col).value = "=" + "+".join(terms)

    # 1. Outer & Inner Sales (supporting 1, 2, 3, or 4 price tiers)
    for idx, ps in enumerate(state.pumps):
        outer_closings = [ps.closing_meter_1, ps.closing_meter_2, ps.closing_meter_3, ps.closing_meter_4]
        inner_closings = [
            ps.inner_closing_meter_1,
            ps.inner_closing_meter_2,
            ps.inner_closing_meter_3,
            ps.inner_closing_meter_4,
        ]
        prices = [ps.price_1, ps.price_2, ps.price_3, ps.price_4]

        lp = (
            next((x for x in lsales.pumps if x.row == ps.row), None)
            or (lsales.pumps[idx] if idx < len(lsales.pumps) else None)
            or lsales.pump(ps.label)
        )
        if lp:
            _write_pump_tier_row(lsales, lp.row, ps.opening_meter, outer_closings, prices)

        if rsales:
            rp = (
                next((x for x in rsales.pumps if x.row == ps.row), None)
                or (rsales.pumps[idx] if idx < len(rsales.pumps) else None)
                or rsales.pump(ps.label)
            )
            if rp:
                _write_pump_tier_row(rsales, rp.row, ps.inner_opening_meter, inner_closings, prices)

    # 2. Tanks & Products Received (Left + mirrored to Right, with dynamic tank expansion)
    for half in layout.halves:
        if not half.tanks:
            continue
        _apply_tank_links(ws, half, state.tanks)
        t = half.tanks
        for idx, ts in enumerate(state.tanks):
            tk = t.tank(ts.label) or (t.tanks[idx] if idx < len(t.tanks) else None)
            if tk is None:
                # Dynamic expansion: if a new tank was added and there is an open row below the tank table
                cand_row = (t.tanks[-1].row + 1) if t.tanks else (t.header_row + 1)
                if cand_row < t.row_limit and ws.cell(row=cand_row, column=t.label_col).value is None:
                    ws.cell(row=cand_row, column=t.label_col).value = ts.label
                    tk = Tank(
                        label=ts.label,
                        row=cand_row,
                        product=ts.product,
                        linked_into=ts.linked_into,
                        has_sales_formula=False,
                    )
                    t.tanks.append(tk)
                else:
                    # Still write to PRODUCTS RECEIVED if matched/allocated
                    rr_only = _ensure_received_row(ws, half, ts.label, idx)
                    if rr_only is not None and half.received:
                        recv_val = 0.0 if ts.under_maintenance else ts.product_received
                        rcell = ws.cell(row=rr_only, column=half.received.value_col)
                        if not is_formula(rcell.value):
                            rcell.value = recv_val if recv_val > 0 else None
                    continue

            o_col, c_col = _physical_dip_pair(ws, t, tk)
            open_dip = ts.opening_dip
            close_dip = ts.opening_dip if ts.under_maintenance else ts.closing_dip
            recv_val = 0.0 if ts.under_maintenance else ts.product_received
            if not is_formula(ws.cell(row=tk.row, column=o_col).value):
                ws.cell(row=tk.row, column=o_col).value = open_dip
            if not is_formula(ws.cell(row=tk.row, column=c_col).value):
                ws.cell(row=tk.row, column=c_col).value = close_dip

            rr = _ensure_received_row(ws, half, ts.label, idx, tk=tk)
            if rr is not None and half.received:
                rcell = ws.cell(row=rr, column=half.received.value_col)
                if not is_formula(rcell.value):
                    rcell.value = recv_val if recv_val > 0 else None

            # Ensure newly added tank rows have their native SALES formula (=Open + Received - Close)
            if t.sales_col and not tk.linked_into:
                scell = ws.cell(row=tk.row, column=t.sales_col)
                if scell.value is None or str(scell.value).strip() == "":
                    eff_o, eff_c = t.dip_pairs[-1]
                    if ws.cell(row=tk.row, column=eff_o).value is None:
                        eff_o, eff_c = o_col, c_col
                    rec_part = (
                        f"+{get_column_letter(half.received.value_col)}{rr}"
                        if (rr is not None and half.received)
                        else ""
                    )
                    scell.value = f"={get_column_letter(eff_o)}{tk.row}{rec_part}-{get_column_letter(eff_c)}{tk.row}"
                    tk.has_sales_formula = True

    # 3. POS & Cash Balancing (Left + mirrored to Right)
    if layout.left.pos:
        lpos = layout.left.pos
        ws.cell(row=lpos.value_row, column=lpos.add_up_col).value = state.add_up
        for col in lpos.all_value_cols:
            c = ws.cell(row=lpos.value_row, column=col)
            if not is_formula(c.value):
                c.value = 0
        for ch in state.pos_channels:
            ws.cell(row=lpos.value_row, column=ch.col).value = ch.amount

    if layout.right and layout.right.pos:
        rpos = layout.right.pos
        if state.mirror_pos_to_inner:
            ws.cell(row=rpos.value_row, column=rpos.add_up_col).value = state.inner_add_up
            for col in rpos.all_value_cols:
                c = ws.cell(row=rpos.value_row, column=col)
                if not is_formula(c.value):
                    c.value = 0
            by_name = {norm(ch.label): ch.amount for ch in state.pos_channels}
            for rlabel, rcol in rpos.channels:
                if norm(rlabel) in by_name:
                    ws.cell(row=rpos.value_row, column=rcol).value = by_name[norm(rlabel)]

    # 4. Expenses & Imprest (Left, optionally mirrored to Right if entered)
    def _write_list(blk, entries: List[ExpenseEntry]):
        if blk is None:
            return
        active = [e for e in entries if e.description.strip() or e.amount != 0]
        r = blk.first_row
        for e in active:
            while r <= blk.last_row and (
                is_formula(ws.cell(row=r, column=blk.desc_col).value)
                or is_formula(ws.cell(row=r, column=blk.amount_col).value)
                or ws.cell(row=r, column=blk.desc_col).value is not None
            ):
                r += 1
            if r > blk.last_row:
                break
            ws.cell(row=r, column=blk.desc_col).value = e.description.strip()
            ws.cell(row=r, column=blk.amount_col).value = e.amount
            r += 1

    _write_list(layout.left.expenses, state.expenses)
    _write_list(layout.left.imprest, state.imprest)

    # 5. Bank lodgment (optional)
    if layout.left.bank:
        for k, r in layout.left.bank.rows.items():
            nk = norm(k)
            if "BANK" in nk and state.bank_name:
                ws.cell(row=r, column=layout.left.bank.value_col).value = state.bank_name
            elif "AMOUNT" in nk and state.bank_amount_deposited > 0:
                ws.cell(row=r, column=layout.left.bank.value_col).value = state.bank_amount_deposited
            elif "STATUS" in nk and state.bank_payment_status:
                ws.cell(row=r, column=layout.left.bank.value_col).value = state.bank_payment_status
            elif "DATE" in nk and state.bank_amount_deposited > 0:
                ws.cell(row=r, column=layout.left.bank.value_col).value = dt_val


def evaluate_summary(ws: Worksheet, layout: SheetLayout, state: StationState) -> StationSummary:
    """Evaluate all key formulas on ``ws`` to produce a live :class:`StationSummary`."""
    ev = SheetEvaluator(ws)
    lsales = layout.left.sales
    total_sales = 0.0
    prod_litres: Dict[str, float] = {}
    for p in lsales.pumps:
        if lsales.amount_col:
            total_sales += ev.try_value(lsales.amount_col, p.row) or 0.0
        lit = sum((ev.try_value(sc, p.row) or 0.0) for sc in lsales.sales_cols)
        prod_litres[p.product or "OTHER"] = prod_litres.get(p.product or "OTHER", 0.0) + lit

    inner_sales = 0.0
    if layout.right and layout.right.sales.amount_col:
        rs = layout.right.sales
        for p in rs.pumps:
            inner_sales += ev.try_value(rs.amount_col, p.row) or 0.0

    total_exp = 0.0
    if layout.left.expenses and layout.left.expenses.total_row:
        total_exp = ev.try_value(layout.left.expenses.amount_col, layout.left.expenses.total_row) or 0.0

    total_pos = sum(ch.amount for ch in state.pos_channels)
    avail_cash = 0.0
    if layout.left.pos and layout.left.pos.available_col:
        avail_cash = ev.try_value(layout.left.pos.available_col, layout.left.pos.value_row) or 0.0

    inner_avail = 0.0
    if layout.right and layout.right.pos and layout.right.pos.available_col:
        inner_avail = ev.try_value(layout.right.pos.available_col, layout.right.pos.value_row) or 0.0

    return StationSummary(
        station_name=state.station_name,
        source_sheet=state.source_sheet,
        target_sheet=state.target_sheet,
        total_sales_amount=round(total_sales, 2),
        inner_total_sales_amount=round(inner_sales, 2),
        product_litres={k: round(v, 2) for k, v in prod_litres.items()},
        total_expenses=round(total_exp, 2),
        total_pos=round(total_pos, 2),
        add_up=round(state.add_up, 2),
        available_cash=round(avail_cash, 2),
        inner_available_cash=round(inner_avail, 2),
        warnings=list(state.warnings),
    )


def rollover_sheet(wb: Workbook, report_date: _dt.date, cached_wb: Optional[Workbook] = None) -> Tuple[Worksheet, SheetLayout]:
    """Duplicate the latest worksheet in ``wb`` and perform the daily rollover."""
    prev_ws = wb.worksheets[-1]
    prev_val = cached_wb[prev_ws.title] if (cached_wb and prev_ws.title in cached_wb.sheetnames) else None
    prev_eval = SheetEvaluator(prev_ws, prev_val)
    target_name = next_sheet_name(wb)

    new_ws = wb.copy_worksheet(prev_ws)
    new_ws.title = target_name
    layout = parse_sheet(new_ws)

    _reset_half_on_rollover(new_ws, prev_ws, prev_eval, layout.left, report_date, roll_pos_addup=True)
    if layout.right:
        _reset_half_on_rollover(new_ws, prev_ws, prev_eval, layout.right, report_date, roll_pos_addup=False)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    return new_ws, layout


def process_workbook(
    source: Union[bytes, str, io.BytesIO],
    state: StationState,
) -> Tuple[bytes, StationSummary]:
    """Duplicate latest sheet, roll over, apply ``state``, and return ``(xlsx_bytes, summary)``."""
    if isinstance(source, (bytes, bytearray)):
        raw = bytes(source)
    elif isinstance(source, io.BytesIO):
        raw = source.getvalue()
    else:
        with open(source, "rb") as f:
            raw = f.read()

    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=False)
    try:
        new_ws, layout = rollover_sheet(wb, state.report_date)
        new_ws.title = state.target_sheet
        apply_station_state(new_ws, layout, state)
        summary = evaluate_summary(new_ws, layout, state)

        out = io.BytesIO()
        wb.save(out)
        data = out.getvalue()
        out.close()
    finally:
        # openpyxl object trees are large (~25-45 MB per workbook); free them immediately
        try:
            wb.close()
        except Exception:
            pass
        del wb
        gc.collect()
    return data, summary
