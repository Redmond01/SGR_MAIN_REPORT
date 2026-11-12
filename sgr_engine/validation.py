"""Pre-flight integrity checks run on every processed workbook before it is queued.

A report is only marked *Done* and added to the download queue when
:func:`validate_output` returns no errors.  Checks performed:

1. The bytes are a valid, uncorrupted ``.xlsx`` (ZIP CRC check).
2. The workbook opens in openpyxl and gained exactly one new sheet.
3. The new sheet is the last sheet and carries the expected ``Sheet{N+1}`` name.
4. The layout anchors (sales / tanks / POS) are still detectable on the new sheet.
5. No formula in the new sheet contains ``#REF!``.
6. Every pump ``AMOUNT`` cell evaluates, and equals the app's own
   ``sum(Sales_t * Price_t)`` across all active price tiers (Left / Outer half).
7. ``AVAILABLE CASH`` evaluates (warning only if it uses unsupported functions).
"""
from __future__ import annotations

import gc
import io
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional

import openpyxl

from .evaluator import SheetEvaluator
from .parser import parse_sheet

__all__ = ["ValidationResult", "validate_output"]

_AMOUNT_TOLERANCE = 0.05  # Naira


@dataclass
class ValidationResult:
    ok: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    sheet_count: int = 0
    size_bytes: int = 0


def validate_output(out_bytes: bytes, state, source_sheet_count: Optional[int] = None) -> ValidationResult:
    """Validate a processed workbook produced by :func:`process_workbook`."""
    errors: List[str] = []
    warnings: List[str] = []
    res = ValidationResult(ok=False, size_bytes=len(out_bytes or b""))

    # 1. Container integrity
    if not out_bytes:
        res.errors.append("Generated workbook is empty (0 bytes).")
        return res
    try:
        with zipfile.ZipFile(io.BytesIO(out_bytes)) as zf:
            bad = zf.testzip()
            if bad is not None:
                res.errors.append(f"Workbook container is corrupted (bad member: {bad}).")
                return res
    except zipfile.BadZipFile as exc:
        res.errors.append(f"Generated file is not a valid .xlsx: {exc}")
        return res

    wb = None
    try:
        # 2. Opens cleanly
        try:
            wb = openpyxl.load_workbook(io.BytesIO(out_bytes), data_only=False)
        except Exception as exc:  # pragma: no cover - defensive
            res.errors.append(f"Workbook could not be re-opened: {exc}")
            return res

        res.sheet_count = len(wb.worksheets)
        if source_sheet_count is not None and res.sheet_count != source_sheet_count + 1:
            errors.append(
                f"Expected {source_sheet_count + 1} sheets after rollover, found {res.sheet_count}."
            )

        # 3. New sheet is last and correctly named
        ws = wb.worksheets[-1]
        if ws.title != state.target_sheet:
            errors.append(f"Last sheet is '{ws.title}', expected '{state.target_sheet}'.")

        # 4. Anchors still detectable
        try:
            layout = parse_sheet(ws)
        except Exception as exc:
            errors.append(f"Layout anchors could not be detected on '{ws.title}': {exc}")
            layout = None

        # 5. No broken references
        ref_cells = []
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and "#REF!" in v.upper():
                    ref_cells.append(c.coordinate)
        if ref_cells:
            shown = ", ".join(ref_cells[:8]) + (" ..." if len(ref_cells) > 8 else "")
            errors.append(f"Broken #REF! formulas in {len(ref_cells)} cell(s): {shown}")

        if layout is not None:
            ev = SheetEvaluator(ws)
            s = layout.left.sales
            tiers = max(1, min(4, int(getattr(state, "price_tiers", 1) or 1)))

            # 6. Pump AMOUNT cells evaluate and match the app's multi-tier calculation
            if s.amount_col:
                for idx, ps in enumerate(state.pumps):
                    lp = (
                        next((x for x in s.pumps if x.row == ps.row), None)
                        or (s.pumps[idx] if idx < len(s.pumps) else None)
                    )
                    if lp is None:
                        warnings.append(f"Pump '{ps.label}' not found on new sheet.")
                        continue
                    val = ev.try_value(s.amount_col, lp.row)
                    if val is None:
                        errors.append(f"AMOUNT for pump '{ps.label}' (row {lp.row}) could not be evaluated.")
                        continue
                    expected = ps.compute_tiers(tiers)["total_amount"]
                    if abs(val - expected) > _AMOUNT_TOLERANCE:
                        errors.append(
                            f"AMOUNT mismatch on pump '{ps.label}': sheet ₦{val:,.2f} vs expected ₦{expected:,.2f}."
                        )
            else:
                warnings.append("No AMOUNT column detected on sales table.")

            # 7. AVAILABLE CASH evaluates
            pos = layout.left.pos
            if pos and pos.available_col:
                if ev.try_value(pos.available_col, pos.value_row) is None:
                    warnings.append("AVAILABLE CASH formula could not be evaluated by the in-app checker (Excel will still calculate it).")
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        del wb
        gc.collect()

    res.errors = errors
    res.warnings = warnings
    res.ok = not errors
    return res
