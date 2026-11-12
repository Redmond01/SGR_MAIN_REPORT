"""Batch template export & import for centralized multi-station entry.

Allows the operator to download one unified Excel workbook covering all 60
stations (with Opening Meters, Prices, Opening Tank Dips, and POS Channels
pre-populated from rollover), fill Closing figures in Excel if desired
(including 1, 2, 3, or 4 price tiers per day), and re-upload it to populate
all 60 station states in one click.
"""
from __future__ import annotations

import io
from typing import Dict, List, Tuple

import openpyxl

from .engine import ExpenseEntry, StationState, TankRowState
from .parser import compact, norm, product_key

__all__ = ["export_batch_template", "import_batch_template"]


def export_batch_template(states: Dict[str, StationState]) -> bytes:
    """Create a multi-sheet Excel template pre-filled with rolled-over openings."""
    wb = openpyxl.Workbook()
    ws_sales = wb.active
    ws_sales.title = "Sales_Pumps"
    ws_sales.append([
        "File_Name",
        "Station",
        "Pump",
        "Product",
        "Price_Tiers (1-4)",
        "Opening_Meter",
        "Closing_Meter_1",
        "Price_1",
        "Closing_Meter_2",
        "Price_2",
        "Closing_Meter_3",
        "Price_3",
        "Closing_Meter_4",
        "Price_4",
        "Inner_Opening_Meter",
        "Inner_Closing_Meter_1",
        "Inner_Closing_Meter_2",
        "Inner_Closing_Meter_3",
        "Inner_Closing_Meter_4",
    ])

    ws_tanks = wb.create_sheet("Tanks_Deliveries")
    ws_tanks.append([
        "File_Name",
        "Station",
        "Tank",
        "Product",
        "Opening_Dip",
        "Closing_Dip",
        "Products_Received",
        "Under_Maintenance",
        "Linked_Into_Tank",
    ])

    ws_pos = wb.create_sheet("POS_Channels")
    ws_pos.append([
        "File_Name",
        "Station",
        "Channel",
        "Amount",
    ])

    ws_exp = wb.create_sheet("Expenses")
    ws_exp.append([
        "File_Name",
        "Station",
        "Type",
        "Description",
        "Amount",
    ])

    for fname, st in states.items():
        tiers = max(1, min(4, int(st.price_tiers)))
        for p in st.pumps:
            ws_sales.append([
                fname,
                st.station_name,
                p.label,
                p.product,
                tiers,
                p.opening_meter,
                p.closing_meter_1,
                p.price_1,
                p.closing_meter_2 if tiers >= 2 else "",
                p.price_2 if tiers >= 2 and p.price_2 > 0 else "",
                p.closing_meter_3 if tiers >= 3 else "",
                p.price_3 if tiers >= 3 and p.price_3 > 0 else "",
                p.closing_meter_4 if tiers >= 4 else "",
                p.price_4 if tiers >= 4 and p.price_4 > 0 else "",
                p.inner_opening_meter,
                p.inner_closing_meter_1,
                p.inner_closing_meter_2 if tiers >= 2 else "",
                p.inner_closing_meter_3 if tiers >= 3 else "",
                p.inner_closing_meter_4 if tiers >= 4 else "",
            ])
        for t in st.tanks:
            ws_tanks.append([
                fname,
                st.station_name,
                t.label,
                t.product,
                t.opening_dip,
                t.closing_dip,
                t.product_received,
                "YES" if t.under_maintenance else "NO",
                t.linked_into or "",
            ])
        for ch in st.pos_channels:
            ws_pos.append([
                fname,
                st.station_name,
                ch.label,
                ch.amount,
            ])
        for ex in st.expenses:
            if ex.description or ex.amount:
                ws_exp.append([fname, st.station_name, "EXPENSE", ex.description, ex.amount])
        for im in st.imprest:
            if im.description or im.amount:
                ws_exp.append([fname, st.station_name, "IMPREST", im.description, im.amount])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _fnum(v, default: float = 0.0) -> float:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").strip())
        except ValueError:
            return default
    return default


def import_batch_template(raw_bytes: bytes, states: Dict[str, StationState]) -> Tuple[int, List[str]]:
    """Populate ``states`` in-place from an uploaded batch template workbook."""
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    updated_stations = set()
    logs: List[str] = []

    by_station_name = {norm(st.station_name): st for st in states.values()}

    def _resolve_state(fname_cell, stname_cell):
        if fname_cell and str(fname_cell) in states:
            return states[str(fname_cell)]
        if stname_cell and norm(stname_cell) in by_station_name:
            return by_station_name[norm(stname_cell)]
        return None

    if "Sales_Pumps" in wb.sheetnames:
        ws = wb["Sales_Pumps"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or len(row) < 8:
                continue
            st = _resolve_state(row[0], row[1])
            if st is None:
                continue
            plabel = compact(row[2])
            p = next((x for x in st.pumps if compact(x.label) == plabel), None)
            if p is None:
                continue
            tiers = max(1, min(4, int(_fnum(row[4], 1))))
            if tiers > st.price_tiers:
                st.set_price_tiers(tiers)
            p.opening_meter = _fnum(row[5], p.opening_meter)
            p.closing_meter_1 = _fnum(row[6], p.closing_meter_1)
            p.price_1 = _fnum(row[7], p.price_1)
            p.closing_meter_2 = _fnum(row[8], p.closing_meter_1) if len(row) > 8 and row[8] not in (None, "") else p.closing_meter_1
            p.price_2 = _fnum(row[9], p.price_1) if len(row) > 9 and row[9] not in (None, "") else (p.price_1 if tiers >= 2 else 0.0)
            p.closing_meter_3 = _fnum(row[10], p.closing_meter_2) if len(row) > 10 and row[10] not in (None, "") else p.closing_meter_2
            p.price_3 = _fnum(row[11], p.price_2) if len(row) > 11 and row[11] not in (None, "") else (p.price_2 if tiers >= 3 else 0.0)
            p.closing_meter_4 = _fnum(row[12], p.closing_meter_3) if len(row) > 12 and row[12] not in (None, "") else p.closing_meter_3
            p.price_4 = _fnum(row[13], p.price_3) if len(row) > 13 and row[13] not in (None, "") else (p.price_3 if tiers >= 4 else 0.0)
            if len(row) > 14 and row[14] not in (None, ""):
                p.inner_opening_meter = _fnum(row[14], p.inner_opening_meter)
            if len(row) > 15 and row[15] not in (None, ""):
                p.inner_closing_meter_1 = _fnum(row[15], p.inner_closing_meter_1)
            p.inner_closing_meter_2 = (
                _fnum(row[16], p.inner_closing_meter_1) if len(row) > 16 and row[16] not in (None, "") else p.inner_closing_meter_1
            )
            p.inner_closing_meter_3 = (
                _fnum(row[17], p.inner_closing_meter_2) if len(row) > 17 and row[17] not in (None, "") else p.inner_closing_meter_2
            )
            p.inner_closing_meter_4 = (
                _fnum(row[18], p.inner_closing_meter_3) if len(row) > 18 and row[18] not in (None, "") else p.inner_closing_meter_3
            )
            updated_stations.add(st.station_name)

    if "Tanks_Deliveries" in wb.sheetnames:
        ws = wb["Tanks_Deliveries"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or len(row) < 7:
                continue
            st = _resolve_state(row[0], row[1])
            if st is None:
                continue
            tlabel = compact(row[2])
            if not tlabel:
                continue
            t = next((x for x in st.tanks if compact(x.label) == tlabel), None)
            if t is None:
                raw_label = str(row[2]).strip()
                prod = str(row[3]).strip().upper() if row[3] else product_key(raw_label)
                next_row = (st.tanks[-1].row + 1) if st.tanks else 20
                t = TankRowState(
                    label=raw_label,
                    product=prod,
                    row=next_row,
                    opening_dip=_fnum(row[4], 0.0),
                    closing_dip=_fnum(row[5], 0.0),
                    product_received=_fnum(row[6], 0.0),
                )
                st.tanks.append(t)
            t.opening_dip = _fnum(row[4], t.opening_dip)
            t.closing_dip = _fnum(row[5], t.closing_dip)
            t.product_received = _fnum(row[6], 0.0)
            if len(row) > 7:
                t.under_maintenance = norm(row[7]) in ("YES", "TRUE", "1", "Y")
            if len(row) > 8:
                link_str = str(row[8]).strip() if row[8] is not None else ""
                t.linked_into = link_str if link_str else None
            updated_stations.add(st.station_name)

    if "POS_Channels" in wb.sheetnames:
        ws = wb["POS_Channels"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or len(row) < 4:
                continue
            st = _resolve_state(row[0], row[1])
            if st is None:
                continue
            ch_name = norm(row[2])
            ch = next((x for x in st.pos_channels if norm(x.label) == ch_name), None)
            if ch is None:
                continue
            ch.amount = _fnum(row[3], 0.0)
            updated_stations.add(st.station_name)

    if "Expenses" in wb.sheetnames:
        ws = wb["Expenses"]
        exp_map: Dict[str, List[ExpenseEntry]] = {}
        imp_map: Dict[str, List[ExpenseEntry]] = {}
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or len(row) < 5:
                continue
            st = _resolve_state(row[0], row[1])
            if st is None:
                continue
            kind = norm(row[2])
            desc = str(row[3]).strip() if row[3] is not None else ""
            amt = _fnum(row[4], 0.0)
            if not desc and amt == 0:
                continue
            if kind == "IMPREST":
                imp_map.setdefault(st.file_name, []).append(ExpenseEntry(desc, amt))
            else:
                exp_map.setdefault(st.file_name, []).append(ExpenseEntry(desc, amt))
            updated_stations.add(st.station_name)
        for fn, items in exp_map.items():
            states[fn].expenses = items
        for fn, items in imp_map.items():
            states[fn].imprest = items

    logs.append(f"Updated {len(updated_stations)} station(s) from batch template.")
    return len(updated_stations), logs
