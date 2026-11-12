"""Print the detected layout for one or more workbooks (latest sheet)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import openpyxl
from openpyxl.utils import get_column_letter as L
from sgr_engine.parser import parse_sheet, LayoutError


def describe(path, sheet=None, brief=False):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet] if sheet else wb.worksheets[-1]
    try:
        lay = parse_sheet(ws)
    except LayoutError as e:
        print(f"!! {os.path.basename(path)}: {e}")
        return None
    name = os.path.basename(path)
    for h in lay.halves:
        s, t, p = h.sales, h.tanks, h.pos
        line = (f"{h.side[0]} sales@{L(s.open_col)}{s.header_row} C={[L(c) for c in s.closing_cols]} "
                f"P={[L(c) for c in s.price_cols]} pumps={len(s.pumps)}")
        if t:
            line += (f" | tank@{L(t.label_col)}{t.header_row} pairs={[(L(a), L(b)) for a, b in t.dip_pairs]} "
                     f"n={len(t.tanks)} links={[(x.label, x.linked_into) for x in t.tanks if x.linked_into]}")
        if p:
            line += f" | pos@{p.header_row} ch={len(p.channels)} addup={L(p.add_up_col)} avail={L(p.available_col) if p.available_col else None}"
        if h.received:
            line += f" | recv={len(h.received.rows)}"
        if h.expenses:
            e = h.expenses
            line += f" | exp {L(e.desc_col)}/{L(e.amount_col)} r{e.first_row}-{e.last_row} tot={e.total_row}"
        line += f" | date={h.date_cells}"
        if h.warnings:
            line += f" | WARN {h.warnings}"
        print(f"{name[:30]:30} {line}")
        if not brief:
            print("     pumps:", [(x.label, x.row) for x in s.pumps])
            if t:
                print("     tanks:", [(x.label, x.row, x.has_sales_formula) for x in t.tanks])
            if p:
                print("     pos:", [c for c, _ in p.channels])
            if h.received:
                print("     recv:", h.received.rows)
    return lay


if __name__ == "__main__":
    brief = "--brief" in sys.argv
    for a in [x for x in sys.argv[1:] if not x.startswith("--")]:
        describe(a, brief=brief)
