"""End-to-end verification across all 60 station workbooks."""
import sys, os, glob, io
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import openpyxl
from sgr_engine.engine import inspect_workbook, process_workbook
from sgr_engine.evaluator import SheetEvaluator
from sgr_engine.parser import parse_sheet

D = "/home/im-redmond/Desktop/Desktop/record - Copy (3)"
files = sorted(glob.glob(D + "/SEPTEMBER*.xlsx"))
print(f"Testing {len(files)} workbooks...")

ok = 0
failures = []
for p in files:
    fname = os.path.basename(p)
    try:
        with open(p, "rb") as f:
            raw = f.read()
        state = inspect_workbook(raw, file_name=fname)
        # Simulate entering +100L on first pump and +50000 on first POS channel
        if state.pumps:
            state.pumps[0].closing_meter_1 = round(state.pumps[0].opening_meter + 100.0, 2)
            state.pumps[0].closing_meter_2 = state.pumps[0].closing_meter_1
            state.pumps[0].inner_closing_meter_1 = round(state.pumps[0].inner_opening_meter + 100.0, 2)
            state.pumps[0].inner_closing_meter_2 = state.pumps[0].inner_closing_meter_1
        if state.tanks:
            state.tanks[0].closing_dip = max(0.0, state.tanks[0].opening_dip - 100.0)
        if state.pos_channels:
            state.pos_channels[0].amount = 50000.0
        out_bytes, summary = process_workbook(raw, state)
        # Verify output workbook opens cleanly and has new sheet
        wb_out = openpyxl.load_workbook(io.BytesIO(out_bytes))
        assert wb_out.worksheets[-1].title == state.target_sheet
        ok += 1
    except Exception as e:
        failures.append((fname, str(e)))
        print(f"FAIL {fname}: {e}")

print(f"RESULT: {ok}/{len(files)} succeeded, {len(failures)} failed.")
