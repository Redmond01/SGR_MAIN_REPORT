"""Streamlit UI for the SGR 60-Station Daily Workbook Automation Engine."""
from __future__ import annotations

import datetime as _dt
import gc
import io
from typing import Dict, List

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from sgr_engine import (
    ExpenseEntry,
    SessionStore,
    StationState,
    StationSummary,
    TankRowState,
    ValidationResult,
    cleanup_stale_sessions,
    export_batch_template,
    import_batch_template,
    inspect_workbook,
    process_workbook,
    product_key,
    validate_output,
)

st.set_page_config(
    page_title="SGR 60-Station Daily Report Engine",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Global CSS for full-width untruncated KPI indexes & bold table typography
st.markdown(
    """
    <style>
    /* Ensure st.metric values and labels never truncate with '...' and always show in full */
    [data-testid="stMetricValue"] {
        font-size: 1.35rem !important;
        font-weight: 800 !important;
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: unset !important;
        line-height: 1.25 !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.9rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        white-space: normal !important;
    }
    /* Custom Full-Amount KPI Index Cards */
    .sgr-kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
        gap: 12px;
        margin: 10px 0 18px 0;
    }
    .sgr-kpi-card {
        background: #ffffff;
        border: 1.5px solid #cbd5e1;
        border-left: 5px solid #0f172a;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
    }
    .sgr-kpi-card.sales-red {
        border-left-color: #d90429;
        background: #fffafa;
    }
    .sgr-kpi-card.pos-blue {
        border-left-color: #1d4ed8;
    }
    .sgr-kpi-card.cash-green {
        border-left-color: #15803d;
        background: #f8fdf9;
    }
    .sgr-kpi-title {
        font-size: 0.82rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #334155;
        margin-bottom: 4px;
    }
    .sgr-kpi-amount {
        font-size: 1.38rem;
        font-weight: 800;
        color: #0f172a;
        white-space: normal;
        word-break: break-word;
        line-height: 1.25;
    }
    .sgr-kpi-amount.red {
        color: #d90429;
    }
    .sgr-kpi-amount.blue {
        color: #1d4ed8;
    }
    .sgr-kpi-amount.green {
        color: #15803d;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

_BOLD_TRANS = str.maketrans(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
    "𝗔𝗕𝗖𝗗𝗘𝗙𝗚𝗛𝗜𝗝𝗞𝗟𝗠𝗡𝗢𝗣𝗤𝗥𝗦𝗧𝗨𝗩𝗪𝗫𝗬𝗭𝗮𝗯𝗰𝗱𝗲𝗳𝗴𝗵𝗶𝗷𝗸𝗹𝗺𝗻𝗼𝗽𝗾𝗿𝘀𝘁𝘂𝘃𝘄𝘅𝘆𝘇𝟬𝟭𝟮𝟯𝟰𝟱𝟲𝟳𝟴𝟵",
)


def _bold_hdr(text: str) -> str:
    """Return a visually bold column header label for Glide Data Grid."""
    return text.translate(_BOLD_TRANS)



# ---------------------------------------------------------------- session state


@st.cache_resource(show_spinner=False)
def _cleanup_once_per_server() -> int:
    """Remove temp folders left behind by browser sessions abandoned for > 24h."""
    return cleanup_stale_sessions(max_age_hours=24.0)


def _init_state():
    # NOTE: workbook bytes are NOT kept in session_state (RAM). Raw uploads and
    # validated outputs live on disk in a per-session folder (see SessionStore).
    defaults = {
        "store_id": None,        # per-session temp folder id
        "states": {},            # file_name -> StationState (small numbers only)
        "saved_snapshots": {},   # file_name -> fingerprint of last saved state
        "summaries": {},         # file_name -> StationSummary
        "editor_rev": {},        # widget_prefix -> int revision counter for immediate onChange refresh
        "selected_file": None,
        "uploader_rev": 0,       # bump to clear the uploader widget (releases its RAM copy)
        "upload_msgs": [],       # messages from the last ingest, shown after the uploader resets
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if not st.session_state["store_id"]:
        _cleanup_once_per_server()
        st.session_state["store_id"] = SessionStore().session_id


_init_state()


def _store() -> SessionStore:
    return SessionStore(st.session_state["store_id"])


def _rev(prefix: str) -> int:
    return st.session_state["editor_rev"].get(prefix, 0)


def _bump_rev(prefix: str):
    st.session_state["editor_rev"][prefix] = _rev(prefix) + 1


# ---------------------------------------------------------------- onChange callbacks for grids


def _on_sales_change(fn: str, tiers: int, key: str, rev_prefix: str):
    """Apply user cell edits to PumpRowState BEFORE the page re-renders so Sales & Amount update onChange."""
    st_state: StationState = st.session_state["states"].get(fn)
    delta = st.session_state.get(key, {})
    edited_rows = delta.get("edited_rows", {}) if isinstance(delta, dict) else {}
    if not st_state or not edited_rows:
        return

    tiers = max(1, min(4, int(tiers)))
    for row_idx_raw, changes in edited_rows.items():
        idx = int(row_idx_raw)
        if idx < 0 or idx >= len(st_state.pumps):
            continue
        p = st_state.pumps[idx]
        prev_calc = p.compute_tiers(tiers)
        old_closings = [tr["closing"] for tr in prev_calc["tiers"]]
        old_sales = [tr["sales"] for tr in prev_calc["tiers"]]
        old_prices = [tr["price"] for tr in prev_calc["tiers"]]
        old_iclosings = [tr["inner_closing"] for tr in prev_calc["tiers"]]
        old_isales = [tr["inner_sales"] for tr in prev_calc["tiers"]]

        if "Opening Meter" in changes and changes["Opening Meter"] is not None:
            p.opening_meter = float(changes["Opening Meter"])
        if "Inner Opening" in changes and changes["Inner Opening"] is not None:
            p.inner_opening_meter = float(changes["Inner Opening"])

        for t in range(1, tiers + 1):
            i = t - 1
            prev_c = p.opening_meter if t == 1 else p.get_closing(t - 1)
            prev_ic = p.inner_opening_meter if t == 1 else p.get_inner_closing(t - 1)
            ck = f"Closing Meter {t}"
            sk = f"Sales {t}"
            pk = f"Price {t} (₦)"
            ick = f"Inner Closing {t}"

            # 1. Outer Closing / Sales for tier t
            if sk in changes and changes[sk] is not None:
                s_val = max(0.0, float(changes[sk]))
                new_c = round(prev_c + s_val, 2)
            elif ck in changes and changes[ck] is not None:
                raw_c = float(changes[ck])
                if raw_c >= prev_c:
                    new_c = round(raw_c, 2)
                elif raw_c > 0:
                    # If user entered a positive number smaller than prev_c in Closing Meter t,
                    # treat it as the tier's sales volume (litres) so Sales t is never zeroed out
                    new_c = round(prev_c + raw_c, 2)
                else:
                    new_c = prev_c
            else:
                if old_closings[i] >= prev_c:
                    new_c = old_closings[i]
                elif old_sales[i] > 0:
                    new_c = round(prev_c + old_sales[i], 2)
                else:
                    new_c = prev_c
            p.set_closing(t, new_c)

            # 2. Price for tier t
            if pk in changes and changes[pk] is not None:
                new_p = max(0.0, float(changes[pk]))
            else:
                fallback_p = p.price_1 if t == 1 else p.get_price(t - 1)
                new_p = old_prices[i] if old_prices[i] > 0 else fallback_p
            p.set_price(t, new_p)

            # 3. Inner Closing for tier t
            if ick in changes and changes[ick] is not None:
                raw_ic = float(changes[ick])
                if raw_ic >= prev_ic:
                    new_ic = round(raw_ic, 2)
                elif raw_ic > 0:
                    new_ic = round(prev_ic + raw_ic, 2)
                else:
                    new_ic = prev_ic
            else:
                if old_iclosings[i] >= prev_ic:
                    new_ic = old_iclosings[i]
                elif old_isales[i] > 0:
                    new_ic = round(prev_ic + old_isales[i], 2)
                else:
                    new_ic = prev_ic
            p.set_inner_closing(t, new_ic)

        # Keep trailing inactive tiers aligned with the last active closing meter
        last_c = p.get_closing(tiers)
        last_ic = p.get_inner_closing(tiers)
        for t in range(tiers + 1, 5):
            p.set_closing(t, last_c)
            p.set_inner_closing(t, last_ic)

    _bump_rev(rev_prefix)


def _on_tanks_change(fn: str, key: str, rev_prefix: str):
    """Apply user cell edits to TankRowState BEFORE re-render so Tank Sales updates onChange."""
    st_state: StationState = st.session_state["states"].get(fn)
    delta = st.session_state.get(key, {})
    edited_rows = delta.get("edited_rows", {}) if isinstance(delta, dict) else {}
    if not st_state or not edited_rows:
        return

    for row_idx_raw, changes in edited_rows.items():
        idx = int(row_idx_raw)
        if idx < 0 or idx >= len(st_state.tanks):
            continue
        t = st_state.tanks[idx]
        if "Opening Dip" in changes and changes["Opening Dip"] is not None:
            t.opening_dip = float(changes["Opening Dip"])
        if "Under Maintenance" in changes and changes["Under Maintenance"] is not None:
            t.under_maintenance = bool(changes["Under Maintenance"])
        if "Closing Dip" in changes and changes["Closing Dip"] is not None:
            t.closing_dip = float(changes["Closing Dip"])
        if "Products Received" in changes and changes["Products Received"] is not None:
            t.product_received = float(changes["Products Received"])
        if "Linked Into (Host Tank)" in changes:
            link_val = str(changes["Linked Into (Host Tank)"] or "").strip()
            t.linked_into = link_val if (link_val and link_val != t.label) else None
        if t.under_maintenance:
            t.closing_dip = t.opening_dip
            t.product_received = 0.0

    _bump_rev(rev_prefix)


def _on_pos_change(fn: str, key: str):
    """Apply POS edits BEFORE re-render so Total POS & Available Cash KPIs update onChange."""
    st_state: StationState = st.session_state["states"].get(fn)
    delta = st.session_state.get(key, {})
    edited_rows = delta.get("edited_rows", {}) if isinstance(delta, dict) else {}
    if not st_state or not edited_rows:
        return
    for row_idx_raw, changes in edited_rows.items():
        idx = int(row_idx_raw)
        if 0 <= idx < len(st_state.pos_channels):
            if "Settlement Amount (₦)" in changes and changes["Settlement Amount (₦)"] is not None:
                st_state.pos_channels[idx].amount = float(changes["Settlement Amount (₦)"])


# ---------------------------------------------------------------- helpers


def _state_fingerprint(s: StationState) -> tuple:
    """Return a hashable snapshot of a station's editable fields to detect unsaved changes."""
    return (
        s.report_date,
        s.price_tiers,
        round(s.add_up, 2),
        s.mirror_pos_to_inner,
        s.bank_name,
        round(s.bank_amount_deposited, 2),
        s.bank_payment_status,
        tuple(
            (
                round(p.opening_meter, 2),
                round(p.closing_meter_1, 2),
                round(p.closing_meter_2, 2),
                round(p.closing_meter_3, 2),
                round(p.closing_meter_4, 2),
                round(p.price_1, 2),
                round(p.price_2, 2),
                round(p.price_3, 2),
                round(p.price_4, 2),
                round(p.inner_opening_meter, 2),
                round(p.inner_closing_meter_1, 2),
                round(p.inner_closing_meter_2, 2),
                round(p.inner_closing_meter_3, 2),
                round(p.inner_closing_meter_4, 2),
            )
            for p in s.pumps
        ),
        tuple(
            (
                t.label,
                round(t.opening_dip, 2),
                round(t.closing_dip, 2),
                round(t.product_received, 2),
                t.under_maintenance,
                t.linked_into or "",
            )
            for t in s.tanks
        ),
        tuple((ch.col, round(ch.amount, 2)) for ch in s.pos_channels),
        tuple((e.description, round(e.amount, 2)) for e in s.expenses),
    )


def _has_unsaved_changes(fn: str) -> bool:
    s = st.session_state["states"].get(fn)
    if s is None:
        return False
    snap = st.session_state["saved_snapshots"].get(fn)
    return _state_fingerprint(s) != snap


def _save_station(fn: str) -> ValidationResult:
    """Process the station, run pre-flight validation, and queue the .xlsx on disk if it passes.

    The station is only marked Done (and added to the Download All queue) when the
    generated workbook passes every integrity check. On failure, nothing is queued
    and any previously queued file for this station is removed (it would be stale).
    """
    s: StationState = st.session_state["states"][fn]
    store = _store()
    if not store.has_raw(fn):
        res = ValidationResult(ok=False, errors=[
            "Original upload is no longer on the server (session expired or server restarted). "
            "Please re-upload this workbook."
        ])
    else:
        try:
            raw = store.load_raw(fn)
            out_bytes, sm = process_workbook(raw, s)
            del raw
            res = validate_output(out_bytes, s, source_sheet_count=s.source_sheet_count or None)
            if res.ok:
                store.save_ready(fn, out_bytes)
                st.session_state["summaries"][fn] = sm
            del out_bytes
        except Exception as exc:  # never let one bad workbook crash the session
            res = ValidationResult(ok=False, errors=[f"Processing failed: {type(exc).__name__}: {exc}"])
        finally:
            gc.collect()

    s.validation_errors = list(res.errors)
    s.validation_warnings = list(res.warnings)
    if res.ok:
        s.is_saved = True
        s.saved_at = _dt.datetime.now().strftime("%H:%M:%S")
        st.session_state["saved_snapshots"][fn] = _state_fingerprint(s)
    else:
        s.is_saved = False
        s.saved_at = None
        store.remove_ready(fn)
    return res


def _is_queued(fn: str) -> bool:
    """A station is in the download queue when saved, unchanged since save, and its file is on disk."""
    s = st.session_state["states"].get(fn)
    return bool(s and s.is_saved and not _has_unsaved_changes(fn) and _store().has_ready(fn))


def _remove_station(fn: str):
    _store().remove_station(fn)
    for k in ("states", "saved_snapshots", "summaries"):
        st.session_state[k].pop(fn, None)
    if st.session_state["selected_file"] == fn:
        st.session_state["selected_file"] = next(iter(st.session_state["states"]), None)


def _clear_all_stations():
    _store().clear()
    for k in ("states", "saved_snapshots", "summaries", "editor_rev"):
        st.session_state[k] = {}
    st.session_state["selected_file"] = None
    st.session_state["uploader_rev"] += 1
    gc.collect()


def _summary_rows(states: Dict[str, StationState], fns: List[str]) -> List[dict]:
    rows = []
    for fn in fns:
        s = states[fn]
        m = _live_preview_metrics(s)
        rows.append({
            "Status": f"✅ Done ({s.saved_at})" if s.is_saved else "⏳ Remaining",
            "Station": s.station_name,
            "Workbook": fn,
            "Rollover": f"{s.source_sheet} ➔ {s.target_sheet}",
            "Date": s.report_date.strftime("%d/%m/%Y"),
            "Price Tiers": s.price_tiers,
            "PMS (L)": round(m["pms_ltr"], 2),
            "AGO (L)": round(m["ago_ltr"], 2),
            "LPG (kg)": round(m["lpg_ltr"], 2),
            "Outer Sales (₦)": round(m["total_sales"], 2),
            "Inner Sales (₦)": round(m["inner_sales"], 2),
            "Total POS (₦)": round(m["total_pos"], 2),
            "Expenses (₦)": round(m["total_exp"], 2),
            "ADD UP (₦)": round(s.add_up, 2),
            "Available Cash (₦)": round(m["avail_cash"], 2),
        })
    return rows


def _make_zip_callable(store: SessionStore, fns: List[str], zip_name: str, summary_rows: List[dict]):
    """Return a zero-arg callable for st.download_button (deferred: runs only when clicked).

    Everything it needs is captured up front because it runs on a separate thread
    with no access to st.session_state.
    """
    fns = list(fns)
    rows = list(summary_rows)

    def _build() -> bytes:
        extra = {}
        if rows:
            buf = io.BytesIO()
            pd.DataFrame(rows).to_excel(buf, index=False, sheet_name="Daily_Summary")
            extra["SGR_Daily_Executive_Summary.xlsx"] = buf.getvalue()
            buf.close()
        zip_path = store.build_zip(fns, zip_name, extra_files=extra)
        with open(zip_path, "rb") as f:
            data = f.read()
        try:
            import os as _os
            _os.remove(zip_path)
        except OSError:
            pass
        return data

    return _build


def _make_file_callable(store: SessionStore, fn: str):
    def _read() -> bytes:
        return store.load_ready(fn)
    return _read


def _render_download_all(states: Dict[str, StationState], key: str, compact_ui: bool = False):
    """Render the 'Download All Completed' button (used in sidebar and Export tab)."""
    queued = [fn for fn in states if _is_queued(fn)]
    stale = [fn for fn, s in states.items() if s.is_saved and _has_unsaved_changes(fn)]
    total = len(states)
    rep_date = next(iter(states.values())).report_date.isoformat() if states else "report"
    zip_name = f"SGR_Station_Reports_{rep_date}_{len(queued)}of{total}.zip"
    if not compact_ui:
        size_mb = sum(_store().ready_size(fn) for fn in queued) / (1024 * 1024)
        st.caption(f"📦 Queue: **{len(queued)} / {total}** validated report(s) ready (~{size_mb:.1f} MB).")
    st.download_button(
        f"📥 Download All Completed ({len(queued)}/{total}) .ZIP",
        data=_make_zip_callable(_store(), queued, zip_name, _summary_rows(states, queued)),
        file_name=zip_name,
        mime="application/zip",
        type="primary",
        use_container_width=True,
        disabled=not queued,
        key=key,
        on_click="ignore",
        help="Packages every saved & validated station workbook into one .ZIP (built only when you click).",
    )
    if stale:
        names = ", ".join(states[fn].station_name for fn in stale[:5]) + (" ..." if len(stale) > 5 else "")
        st.warning(f"🟡 {len(stale)} station(s) edited after saving are NOT in the ZIP until re-saved: {names}")


def _ingest_files(file_pairs: List[tuple]) -> List[str]:
    """Parse and roll over newly uploaded ``(file_name, raw_bytes)`` pairs; raw bytes go to disk."""
    msgs: List[str] = []
    store = _store()
    prog = st.progress(0.0, text="Inspecting workbooks & detecting dynamic anchors...")
    total = len(file_pairs)
    for idx, (fname, raw) in enumerate(file_pairs):
        if fname in st.session_state["states"]:
            msgs.append(f"ℹ️ `{fname}` is already loaded — skipped (remove it first to replace it).")
        else:
            try:
                st_state = inspect_workbook(raw, file_name=fname)
                store.save_raw(fname, raw)
                st.session_state["states"][fname] = st_state
                st.session_state["saved_snapshots"][fname] = _state_fingerprint(st_state)
            except Exception as exc:
                msgs.append(f"❌ Could not parse `{fname}`: {exc}")
        gc.collect()
        prog.progress((idx + 1) / max(1, total), text=f"Parsed {idx + 1}/{total}: {fname}")
    prog.empty()
    if st.session_state["states"] and not st.session_state["selected_file"]:
        st.session_state["selected_file"] = next(iter(st.session_state["states"].keys()))
    return msgs


def _live_preview_metrics(st_state: StationState) -> Dict[str, float]:
    """Fast in-UI preview of Sales, Litres, POS, Expenses, and Available Cash across 1..4 price tiers."""
    total_sales = 0.0
    inner_sales = 0.0
    pms_ltr = 0.0
    ago_ltr = 0.0
    lpg_ltr = 0.0
    tiers = max(1, min(4, int(st_state.price_tiers)))
    tier_amounts = [0.0] * tiers
    for p in st_state.pumps:
        calc = p.compute_tiers(tiers)
        lit = calc["total_sales_ltr"]
        amt = calc["total_amount"]
        iamt = calc["inner_total_amount"]
        for i, tr in enumerate(calc["tiers"]):
            tier_amounts[i] += tr["amount"]

        total_sales += amt
        inner_sales += iamt
        if p.product == "PMS":
            pms_ltr += lit
        elif p.product == "AGO":
            ago_ltr += lit
        elif p.product == "LPG":
            lpg_ltr += lit

    total_pos = sum(ch.amount for ch in st_state.pos_channels)
    total_exp = sum(e.amount for e in st_state.expenses)
    avail_cash = total_sales - total_exp - total_pos + st_state.add_up
    return {
        "total_sales": round(total_sales, 2),
        "inner_sales": round(inner_sales, 2),
        "pms_ltr": round(pms_ltr, 2),
        "ago_ltr": round(ago_ltr, 2),
        "lpg_ltr": round(lpg_ltr, 2),
        "total_pos": round(total_pos, 2),
        "total_exp": round(total_exp, 2),
        "avail_cash": round(avail_cash, 2),
        "tier_amounts": [round(x, 2) for x in tier_amounts],
    }


# ---------------------------------------------------------------- sidebar


with st.sidebar:
    st.title("⛽ SGR Automation")
    st.caption("Dynamic Anchor Ingestion & Rollover Engine")

    uploaded = st.file_uploader(
        "1. Batch Upload Station Workbooks (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        help=(
            "Select one or more station workbooks. They are stored on the server's disk for this session, "
            "and the upload box clears itself afterwards to free memory. You can upload more files at any "
            "time — stations already loaded stay loaded."
        ),
        key=f"uploader_{st.session_state['uploader_rev']}",
    )
    if uploaded:
        pairs = [(u.name, u.getvalue()) for u in uploaded]
        st.session_state["upload_msgs"] = _ingest_files(pairs)
        del pairs, uploaded
        # Reset the uploader widget so Streamlit releases its in-memory copy of the files
        st.session_state["uploader_rev"] += 1
        gc.collect()
        st.rerun()
    for _m in st.session_state.pop("upload_msgs", []):
        (st.error if _m.startswith("❌") else st.info)(_m)

    states: Dict[str, StationState] = st.session_state["states"]
    if states:
        total_cnt = len(states)
        done_list = [fn for fn, s in states.items() if s.is_saved]
        rem_list = [fn for fn, s in states.items() if not s.is_saved]
        done_cnt = len(done_list)
        rem_cnt = len(rem_list)

        st.divider()
        st.subheader(f"📊 Reports Progress: {done_cnt} / {total_cnt}")
        st.progress(done_cnt / max(1, total_cnt), text=f"✅ {done_cnt} Done  |  ⏳ {rem_cnt} Remaining")
        _render_download_all(states, key="dl_all_sidebar", compact_ui=True)

        filter_mode = st.radio(
            "Filter Stations",
            options=["ALL", "REMAINING", "DONE"],
            format_func=lambda m: (
                f"All ({total_cnt})"
                if m == "ALL"
                else (f"⏳ Remaining ({rem_cnt})" if m == "REMAINING" else f"✅ Done ({done_cnt})")
            ),
            horizontal=True,
            label_visibility="collapsed",
        )
        search_q = st.text_input("🔍 Search station...", value="", placeholder="Type station name...", label_visibility="collapsed")

        if filter_mode == "REMAINING":
            visible_fns = rem_list
        elif filter_mode == "DONE":
            visible_fns = done_list
        else:
            visible_fns = list(states.keys())

        if search_q.strip():
            q = search_q.strip().lower()
            visible_fns = [fn for fn in visible_fns if q in states[fn].station_name.lower() or q in fn.lower()]

        st.caption("Click any station below to work on it (reports arrive at different times):")
        with st.container(height=280):
            if not visible_fns:
                st.caption("No stations match this filter.")
            for fn in visible_fns:
                s = states[fn]
                is_current = fn == st.session_state["selected_file"]
                dirty = _has_unsaved_changes(fn)
                if dirty:
                    icon = "🟡"
                    status_tag = " [Unsaved]"
                elif s.is_saved:
                    icon = "✅"
                    status_tag = ""
                else:
                    icon = "⏳"
                    status_tag = ""
                pc_tag = f" ⚡{s.price_tiers}P" if s.price_tiers > 1 else ""
                btn_label = f"{icon} {s.station_name}{pc_tag}{status_tag}"
                if st.button(
                    btn_label,
                    key=f"side_st_{fn}",
                    type="primary" if is_current else "secondary",
                    use_container_width=True,
                ):
                    st.session_state["selected_file"] = fn
                    st.rerun()

        with st.expander("🗑️ Remove Stations / Start a New Day"):
            st.caption(
                "Removing the files from the upload box no longer wipes your work. "
                "Use these buttons to remove a station or clear everything."
            )
            if st.session_state["selected_file"] in states:
                cur_name = states[st.session_state["selected_file"]].station_name
                if st.button(f"Remove '{cur_name}' only", use_container_width=True):
                    _remove_station(st.session_state["selected_file"])
                    st.rerun()
            confirm_clear = st.checkbox(
                "I have downloaded everything I need",
                key=f"confirm_clear_all_{st.session_state['uploader_rev']}",
            )
            if st.button(
                "🧹 Clear ALL Stations & Files",
                use_container_width=True,
                disabled=not confirm_clear,
            ):
                _clear_all_stations()
                st.rerun()

        st.divider()
        st.subheader("⚙️ Global Date & Price Controls")
        first_state = next(iter(states.values()))
        global_date = st.date_input("Report Date (B1 / AD1)", value=first_state.report_date)
        if st.button("📅 Apply Date to All Stations", use_container_width=True):
            for fn, st_s in states.items():
                st_s.report_date = global_date
                if not st_s.is_saved:
                    st.session_state["saved_snapshots"][fn] = _state_fingerprint(st_s)
            st.toast(f"Applied {global_date.strftime('%d/%m/%Y')} to all {len(states)} stations.")

        with st.expander("💰 Global Product Price Override"):
            c1, c2, c3 = st.columns(3)
            g_pms = c1.number_input("PMS ₦/L", value=0.0, step=1.0, format="%.0f")
            g_ago = c2.number_input("AGO ₦/L", value=0.0, step=1.0, format="%.0f")
            g_lpg = c3.number_input("LPG ₦/kg", value=0.0, step=1.0, format="%.0f")
            if st.button("Apply Global Price 1 to All Stations", use_container_width=True):
                for fn, st_s in states.items():
                    for p in st_s.pumps:
                        if p.product == "PMS" and g_pms > 0:
                            p.price_1 = g_pms
                        elif p.product == "AGO" and g_ago > 0:
                            p.price_1 = g_ago
                        elif p.product == "LPG" and g_lpg > 0:
                            p.price_1 = g_lpg
                    _bump_rev(f"sales_{fn}")
                st.toast("Updated Price 1 across all stations.")

        with st.expander("⚡ Multi-Tier Price Change Selector (2, 3, or 4 Prices)"):
            st.caption("Select stations that experienced 1, 2, or 3+ price changes today:")
            current_pc = [fn for fn, s in states.items() if s.price_tiers > 1]
            selected_pc = st.multiselect(
                "Stations with Price Change",
                options=list(states.keys()),
                default=current_pc,
                format_func=lambda fn: states[fn].station_name,
            )
            num_tiers_batch = st.selectbox(
                "Number of Active Prices Today (Tiers)",
                options=[2, 3, 4],
                format_func=lambda n: f"{n} Prices ({n - 1} Price Change{'s' if n > 2 else ''})",
                index=0,
            )
            st.markdown("**Batch Price Override for Selected Stations (Optional)**")
            cp1, cp2 = st.columns(2)
            new_pms_p2 = cp1.number_input("PMS Price 2 (₦)", value=0.0, step=1.0, format="%.0f")
            new_ago_p2 = cp2.number_input("AGO Price 2 (₦)", value=0.0, step=1.0, format="%.0f")
            new_pms_p3 = new_ago_p3 = new_pms_p4 = new_ago_p4 = 0.0
            if num_tiers_batch >= 3:
                cp3, cp4 = st.columns(2)
                new_pms_p3 = cp3.number_input("PMS Price 3 (₦)", value=0.0, step=1.0, format="%.0f")
                new_ago_p3 = cp4.number_input("AGO Price 3 (₦)", value=0.0, step=1.0, format="%.0f")
            if num_tiers_batch >= 4:
                cp5, cp6 = st.columns(2)
                new_pms_p4 = cp5.number_input("PMS Price 4 (₦)", value=0.0, step=1.0, format="%.0f")
                new_ago_p4 = cp6.number_input("AGO Price 4 (₦)", value=0.0, step=1.0, format="%.0f")

            if st.button("Apply Tier Count & Prices to Selected Stations", use_container_width=True):
                pc_set = set(selected_pc)
                for fn, s in states.items():
                    if fn in pc_set:
                        s.set_price_tiers(num_tiers_batch)
                        st.session_state[f"pc_tiers_{fn}"] = s.price_tiers
                        for p in s.pumps:
                            if p.product == "PMS":
                                if new_pms_p2 > 0:
                                    p.price_2 = new_pms_p2
                                if new_pms_p3 > 0:
                                    p.price_3 = new_pms_p3
                                if new_pms_p4 > 0:
                                    p.price_4 = new_pms_p4
                            elif p.product == "AGO":
                                if new_ago_p2 > 0:
                                    p.price_2 = new_ago_p2
                                if new_ago_p3 > 0:
                                    p.price_3 = new_ago_p3
                                if new_ago_p4 > 0:
                                    p.price_4 = new_ago_p4
                    else:
                        s.set_price_tiers(1)
                        st.session_state[f"pc_tiers_{fn}"] = 1
                    for t_idx in range(1, 5):
                        _bump_rev(f"sales_{fn}_{t_idx}")
                st.toast(f"Configured {num_tiers_batch} price tiers on {len(selected_pc)} station(s).")


# ---------------------------------------------------------------- main area


st.header("⛽ SGR Retail Fuel Station Daily Workbook Automation Engine")

states = st.session_state["states"]
if not states:
    st.info(
        "👈 **Upload your monthly station workbooks (`.xlsx`)** in the sidebar "
        "to automatically inspect anchor coordinates, roll over opening meters/dippings/ADD UP, and begin daily entry."
    )
    st.stop()

done_count = sum(1 for s in states.values() if s.is_saved)
rem_count = len(states) - done_count
pc_count = sum(1 for s in states.values() if s.price_tiers > 1)
m1, m2, m3, m4 = st.columns(4)
m1.metric("Reports Completed", f"{done_count} / {len(states)}")
m2.metric("Reports Remaining", f"{rem_count}")
m3.metric("Price Change Stations", f"{pc_count}")
m4.metric("Target Report Date", next(iter(states.values())).report_date.strftime("%d %b %Y"))

tab_station, tab_batch, tab_export = st.tabs([
    "🏪 1. Station Data-Entry Workspace",
    "📋 2. Multi-Station Batch Grid & Excel Import",
    "📦 3. Executive Summary & .ZIP Download",
])


# ================================================================ TAB 1: STATION WORKSPACE
with tab_station:
    file_keys = list(states.keys())
    if st.session_state["selected_file"] not in states:
        st.session_state["selected_file"] = file_keys[0]

    cur_fn = st.session_state["selected_file"]
    st_state = states[cur_fn]
    cur_idx = file_keys.index(cur_fn)

    # Top Station Header & Price Tiers Selector
    hdr1, hdr2, hdr3 = st.columns([3, 2, 2])
    status_badge = (
        f"✅ Saved at {st_state.saved_at}"
        if (st_state.is_saved and not _has_unsaved_changes(cur_fn))
        else ("🟡 Unsaved Changes" if _has_unsaved_changes(cur_fn) else "⏳ Pending (Not Yet Saved)")
    )
    hdr1.markdown(
        f"### {st_state.station_name} &nbsp; `{st_state.source_sheet} ➔ {st_state.target_sheet}`\n"
        f"**Status:** {status_badge}"
    )

    st_state.report_date = hdr2.date_input(
        "Station Report Date",
        value=st_state.report_date,
        key=f"st_date_{cur_fn}",
    )

    tier_labels = {
        1: "1 Price (Standard Day)",
        2: "⚡ 2 Prices (1 Change)",
        3: "⚡ 3 Prices (2 Changes)",
        4: "⚡ 4 Prices (3 Changes)",
    }
    chosen_tiers = hdr3.selectbox(
        "Active Price Tiers Today",
        options=[1, 2, 3, 4],
        index=max(0, min(3, int(st_state.price_tiers) - 1)),
        format_func=lambda t: tier_labels[t],
        key=f"pc_tiers_{st_state.file_name}",
    )
    st_state.set_price_tiers(chosen_tiers)

    # Live KPI ribbon at top of station workspace (updates immediately onChange, always shows full amount)
    kpis = _live_preview_metrics(st_state)
    lpg_part = f" &nbsp;|&nbsp; LPG: {kpis['lpg_ltr']:,.2f} kg" if kpis["lpg_ltr"] > 0 else ""
    st.markdown(
        f"""
        <div class="sgr-kpi-grid">
            <div class="sgr-kpi-card sales-red">
                <div class="sgr-kpi-title">Outer Sales (₦) — Total Amount</div>
                <div class="sgr-kpi-amount red">₦{kpis['total_sales']:,.2f}</div>
            </div>
            <div class="sgr-kpi-card pos-blue">
                <div class="sgr-kpi-title">POS Sales (₦) — Total Amount</div>
                <div class="sgr-kpi-amount blue">₦{kpis['total_pos']:,.2f}</div>
            </div>
            <div class="sgr-kpi-card cash-green">
                <div class="sgr-kpi-title">Available Cash (₦) — Full Balance</div>
                <div class="sgr-kpi-amount green">₦{kpis['avail_cash']:,.2f}</div>
            </div>
            <div class="sgr-kpi-card">
                <div class="sgr-kpi-title">Total Volume (PMS / AGO Litres)</div>
                <div class="sgr-kpi-amount">PMS: {kpis['pms_ltr']:,.2f} L &nbsp;|&nbsp; AGO: {kpis['ago_ltr']:,.2f} L{lpg_part}</div>
            </div>
            <div class="sgr-kpi-card">
                <div class="sgr-kpi-title">Expenses (₦) — Total Amount</div>
                <div class="sgr-kpi-amount">₦{kpis['total_exp']:,.2f}</div>
            </div>
            <div class="sgr-kpi-card">
                <div class="sgr-kpi-title">Rolled ADD UP (₦) — Full Opening Cash</div>
                <div class="sgr-kpi-amount">₦{st_state.add_up:,.2f}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------------- A. Sales & Inner Meter Grid (Immediate onChange calculation)
    tiers = st_state.price_tiers
    st.subheader(f"A. Sales & Inner Meter Readings — {st_state.station_name}")
    if tiers > 1:
        tier_Formula_str = " + ".join(f"(Sales {t} × Price {t})" for t in range(1, tiers + 1))
        tier_sum_parts = " &nbsp;+&nbsp; ".join(
            f"**Tier {t}:** ₦{kpis['tier_amounts'][t - 1]:,.2f}" for t in range(1, tiers + 1)
        )
        st.caption(
            f"Type a **Closing Meter**, **Sales**, or **Price** and press **Enter** / **Tab**: "
            f"`Amount (₦) = {tier_Formula_str}` updates immediately on change.  \n"
            f"📊 **Live Multi-Tier Sum:** {tier_sum_parts} &nbsp;=&nbsp; **Total Outer Sales: ₦{kpis['total_sales']:,.2f}**"
        )
    else:
        st.caption(
            "Type a **Closing Meter**, **Sales**, or **Price** and press **Enter** / **Tab**: `Sales` and `Amount (₦)` calculate immediately on change."
        )

    sales_rows = []
    disabled_cols = ["Pump", "Product", "Total Sales Ltrs", "Amount (₦)", "Inner Sales Ltrs", "Inner Amount (₦)"]

    for p in st_state.pumps:
        calc = p.compute_tiers(tiers)
        row_dict = {
            "Pump": p.label,
            "Product": p.product,
            "Opening Meter": p.opening_meter,
        }
        for tr in calc["tiers"]:
            t_idx = tr["tier"]
            row_dict[f"Closing Meter {t_idx}"] = tr["closing"]
            row_dict[f"Sales {t_idx}"] = tr["sales"]
            row_dict[f"Price {t_idx} (₦)"] = tr["price"]

        row_dict["Total Sales Ltrs"] = calc["total_sales_ltr"]
        row_dict["Amount (₦)"] = calc["total_amount"]
        row_dict["Inner Opening"] = p.inner_opening_meter

        for tr in calc["tiers"]:
            t_idx = tr["tier"]
            row_dict[f"Inner Closing {t_idx}"] = tr["inner_closing"]

        row_dict["Inner Sales Ltrs"] = calc["inner_total_ltr"]
        row_dict["Inner Amount (₦)"] = calc["inner_total_amount"]
        sales_rows.append(row_dict)

    sales_df = pd.DataFrame(sales_rows)
    sales_rev_prefix = f"sales_{cur_fn}_{tiers}"
    sales_ed_key = f"{sales_rev_prefix}_v{_rev(sales_rev_prefix)}"

    # Style [Pump, Product, Opening, Closing] in BOLD, and [Sales, Amount] in BOLD RED
    sales_bold_cols = (
        ["Pump", "Product", "Opening Meter"]
        + [f"Closing Meter {t}" for t in range(1, tiers + 1)]
        + [f"Price {t} (₦)" for t in range(1, tiers + 1)]
        + ["Inner Opening"]
        + [f"Inner Closing {t}" for t in range(1, tiers + 1)]
    )
    sales_red_cols = (
        [f"Sales {t}" for t in range(1, tiers + 1)]
        + ["Total Sales Ltrs", "Amount (₦)", "Inner Sales Ltrs", "Inner Amount (₦)"]
    )
    styled_sales_df = (
        sales_df.style
        .set_properties(subset=sales_bold_cols, **{"font-weight": "bold", "color": "#0f172a"})
        .set_properties(
            subset=sales_red_cols,
            **{"color": "#d90429", "font-weight": "bold", "background-color": "#fff5f5"},
        )
    )

    sales_col_cfg = {
        "Pump": st.column_config.TextColumn(_bold_hdr("Pump"), pinned=True),
        "Product": st.column_config.TextColumn(_bold_hdr("Product"), pinned=True),
        "Opening Meter": st.column_config.NumberColumn(_bold_hdr("Opening Meter"), format="%.2f"),
        "Total Sales Ltrs": st.column_config.NumberColumn(_bold_hdr("Total Sales Ltrs"), format="%.2f"),
        "Amount (₦)": st.column_config.NumberColumn(_bold_hdr("Amount (₦)"), format="₦%.2f"),
        "Inner Opening": st.column_config.NumberColumn(_bold_hdr("Inner Opening"), format="%.2f"),
        "Inner Sales Ltrs": st.column_config.NumberColumn(_bold_hdr("Inner Sales Ltrs"), format="%.2f"),
        "Inner Amount (₦)": st.column_config.NumberColumn(_bold_hdr("Inner Amount (₦)"), format="₦%.2f"),
    }
    for t in range(1, tiers + 1):
        sales_col_cfg[f"Closing Meter {t}"] = st.column_config.NumberColumn(
            _bold_hdr(f"Closing Meter {t}"), format="%.2f"
        )
        sales_col_cfg[f"Sales {t}"] = st.column_config.NumberColumn(
            _bold_hdr(f"Sales {t}"), format="%.2f"
        )
        sales_col_cfg[f"Price {t} (₦)"] = st.column_config.NumberColumn(
            _bold_hdr(f"Price {t} (₦)"), format="₦%.2f"
        )
        sales_col_cfg[f"Inner Closing {t}"] = st.column_config.NumberColumn(
            _bold_hdr(f"Inner Closing {t}"), format="%.2f"
        )

    st.data_editor(
        styled_sales_df,
        key=sales_ed_key,
        disabled=disabled_cols,
        on_change=_on_sales_change,
        args=(cur_fn, tiers, sales_ed_key, sales_rev_prefix),
        column_config=sales_col_cfg,
        hide_index=True,
        use_container_width=True,
    )

    # ---------------- B. Tanks & Deliveries + Dynamic Manifolding (Immediate onChange calculation)
    col_tanks, col_pos = st.columns([3, 2])
    with col_tanks:
        st.subheader(f"B. Tank Dippings, Products Received & Manifolding ({len(st_state.tanks)} Tanks)")
        st.caption(
            "Dynamically detects all station tanks & `PRODUCTS RECEIVED` rows (automatically picks up newly added tanks). "
            "Mirrored to the Right (Inner Meter) half, and `Tank Sales` calculates immediately `onChange`."
        )
        tank_options = [""] + [t.label for t in st_state.tanks]
        tanks_df = pd.DataFrame([
            {
                "Tank": t.label,
                "Product": t.product,
                "Opening Dip": t.opening_dip,
                "Closing Dip": t.opening_dip if t.under_maintenance else t.closing_dip,
                "Products Received": 0.0 if t.under_maintenance else t.product_received,
                "Tank Sales": 0.0 if t.under_maintenance else round(t.opening_dip + t.product_received - t.closing_dip, 2),
                "Under Maintenance": t.under_maintenance,
                "Linked Into (Host Tank)": t.linked_into or "",
            }
            for t in st_state.tanks
        ])
        styled_tanks_df = (
            tanks_df.style
            .set_properties(
                subset=["Tank", "Product", "Opening Dip", "Closing Dip", "Products Received"],
                **{"font-weight": "bold", "color": "#0f172a"},
            )
            .set_properties(
                subset=["Tank Sales"],
                **{"color": "#d90429", "font-weight": "bold", "background-color": "#fff5f5"},
            )
        )
        tanks_rev_prefix = f"tanks_{cur_fn}"
        tanks_ed_key = f"{tanks_rev_prefix}_v{_rev(tanks_rev_prefix)}"
        st.data_editor(
            styled_tanks_df,
            key=tanks_ed_key,
            disabled=["Tank", "Product", "Tank Sales"],
            on_change=_on_tanks_change,
            args=(cur_fn, tanks_ed_key, tanks_rev_prefix),
            column_config={
                "Tank": st.column_config.TextColumn(_bold_hdr("Tank"), pinned=True),
                "Product": st.column_config.TextColumn(_bold_hdr("Product"), pinned=True),
                "Opening Dip": st.column_config.NumberColumn(_bold_hdr("Opening Dip"), format="%.2f"),
                "Closing Dip": st.column_config.NumberColumn(_bold_hdr("Closing Dip"), format="%.2f"),
                "Products Received": st.column_config.NumberColumn(_bold_hdr("Products Received"), format="%.2f"),
                "Tank Sales": st.column_config.NumberColumn(_bold_hdr("Tank Sales"), format="%.2f"),
                "Under Maintenance": st.column_config.CheckboxColumn(_bold_hdr("Under Maintenance")),
                "Linked Into (Host Tank)": st.column_config.SelectboxColumn(
                    _bold_hdr("Linked Into (Host Tank)"),
                    options=tank_options,
                    help="If this tank is manifolded into another tank today, select the host tank.",
                ),
            },
            hide_index=True,
            use_container_width=True,
        )

        tot_open_dip = sum(t.opening_dip for t in st_state.tanks)
        tot_recv = sum(0.0 if t.under_maintenance else t.product_received for t in st_state.tanks)
        tot_close_dip = sum(t.opening_dip if t.under_maintenance else t.closing_dip for t in st_state.tanks)
        tot_tank_sales = sum(
            0.0 if t.under_maintenance else (t.opening_dip + t.product_received - t.closing_dip)
            for t in st_state.tanks
        )
        st.markdown(
            f"**Total Opening Dip:** `{tot_open_dip:,.2f} L` &nbsp;|&nbsp; "
            f"**Total Products Received:** `{tot_recv:,.2f} L` &nbsp;|&nbsp; "
            f"**Total Closing Dip:** `{tot_close_dip:,.2f} L` &nbsp;|&nbsp; "
            f"<span style='color:#d90429;font-weight:800;'>Total Tank Sales: {tot_tank_sales:,.2f} L</span>",
            unsafe_allow_html=True,
        )

        with st.expander("➕ Add New Tank (If Station Increased in Tanks)"):
            st.caption(
                "If this station added a tank in its Excel sheet, it is detected automatically on upload. "
                "You can also add a newly commissioned tank directly here to include it in both **Tank Dippings** and **Products Received**."
            )
            nt1, nt2, nt3, nt4 = st.columns([1.4, 1.2, 1.2, 1.2])
            new_tk_label = nt1.text_input("New Tank Label", value="", placeholder="e.g. PMS 4 or AGO 2", key=f"new_tk_lbl_{cur_fn}")
            new_tk_open = nt2.number_input("Opening Dip", value=0.0, step=100.0, format="%.2f", key=f"new_tk_op_{cur_fn}")
            new_tk_recv = nt3.number_input("Product Received", value=0.0, step=100.0, format="%.2f", key=f"new_tk_rc_{cur_fn}")
            new_tk_close = nt4.number_input("Closing Dip", value=0.0, step=100.0, format="%.2f", key=f"new_tk_cl_{cur_fn}")
            if st.button("➕ Append Tank to Station", key=f"btn_add_tk_{cur_fn}", use_container_width=True):
                clean_lbl = new_tk_label.strip().upper()
                if not clean_lbl:
                    st.warning("Please enter a Tank Label (e.g. `PMS 4` or `AGO 2`).")
                else:
                    prod = product_key(clean_lbl) or "PMS"
                    next_r = (st_state.tanks[-1].row + 1) if st_state.tanks else 20
                    st_state.tanks.append(
                        TankRowState(
                            label=clean_lbl,
                            product=prod,
                            row=next_r,
                            opening_dip=float(new_tk_open),
                            closing_dip=float(new_tk_close),
                            product_received=float(new_tk_recv),
                        )
                    )
                    _bump_rev(tanks_rev_prefix)
                    st.toast(f"Added tank {clean_lbl} to {st_state.station_name}!")
                    st.rerun()

    # ---------------- C. POS & Cash Balancing
    with col_pos:
        st.subheader("C. POS Terminals & Cash Balancing")
        st.caption("All POS channels reset to `0`. Today's `ADD UP` = yesterday's `AVAILABLE CASH`.")
        ac1, ac2 = st.columns(2)
        st_state.add_up = ac1.number_input(
            "Today's ADD UP (₦)",
            value=float(st_state.add_up),
            step=100.0,
            format="%.2f",
            key=f"addup_{st_state.file_name}",
        )
        st_state.mirror_pos_to_inner = ac2.checkbox(
            "Mirror POS to Inner Half",
            value=st_state.mirror_pos_to_inner,
            key=f"mirror_pos_{st_state.file_name}",
        )
        pos_df = pd.DataFrame([
            {"POS / Channel": ch.label, "Settlement Amount (₦)": ch.amount}
            for ch in st_state.pos_channels
        ])
        styled_pos_df = pos_df.style.set_properties(
            subset=["POS / Channel", "Settlement Amount (₦)"],
            **{"font-weight": "bold", "color": "#0f172a"},
        )
        pos_ed_key = f"pos_ed_{cur_fn}"
        st.data_editor(
            styled_pos_df,
            key=pos_ed_key,
            disabled=["POS / Channel"],
            on_change=_on_pos_change,
            args=(cur_fn, pos_ed_key),
            column_config={
                "POS / Channel": st.column_config.TextColumn(_bold_hdr("POS / Channel")),
                "Settlement Amount (₦)": st.column_config.NumberColumn(
                    _bold_hdr("Settlement Amount (₦)"), format="₦%.2f"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )

    # Ensure Glide Data Grid canvas renders all headers and editable Opening/Closing cells in bold
    components.html(
        """
        <script>
        (function() {
            const pWin = window.parent;
            if (!pWin || !pWin.CanvasRenderingContext2D) return;
            const proto = pWin.CanvasRenderingContext2D.prototype;
            if (!proto._sgrBoldPatched) {
                const origFillText = proto.fillText;
                proto.fillText = function(text, x, y, maxWidth) {
                    const prevFont = this.font || "";
                    const prevFill = this.fillStyle;
                    if (this.canvas && this.canvas.closest && this.canvas.closest('[data-testid="stDataFrame"]')) {
                        if (prevFont.includes("400 ")) {
                            this.font = prevFont.replace("400 ", "700 ");
                        } else if (!prevFont.includes("bold") && !prevFont.includes("700") && !prevFont.includes("600")) {
                            this.font = "700 " + prevFont;
                        }
                        if (typeof y === "number" && y <= 36) {
                            this.fillStyle = "#0f172a";
                        }
                    }
                    const res = maxWidth !== undefined
                        ? origFillText.call(this, text, x, y, maxWidth)
                        : origFillText.call(this, text, x, y);
                    this.font = prevFont;
                    this.fillStyle = prevFill;
                    return res;
                };
                proto._sgrBoldPatched = true;
            }
            const canvases = pWin.document.querySelectorAll('[data-testid="stDataFrame"] canvas');
            canvases.forEach(c => {
                c.dispatchEvent(new pWin.MouseEvent("mousemove", { bubbles: true, clientX: 1, clientY: 1 }));
            });
        })();
        </script>
        """,
        height=0,
    )

    # ---------------- D. Optional Expenses & Bank Lodgment
    with st.expander("🧾 D. Optional Expenses, Imprest & Bank Lodgment", expanded=bool(st_state.expenses)):
        ec1, ec2 = st.columns(2)
        with ec1:
            st.markdown("**Station Expenses**")
            exp_rows = st_state.expenses if st_state.expenses else [ExpenseEntry("", 0.0)]
            exp_df = pd.DataFrame([{"Description": e.description, "Amount (₦)": e.amount} for e in exp_rows])
            edited_exp = st.data_editor(
                exp_df,
                key=f"exp_ed_{st_state.file_name}",
                num_rows="dynamic",
                use_container_width=True,
            )
            st_state.expenses = [
                ExpenseEntry(str(r.get("Description") or "").strip(), float(r.get("Amount (₦)") or 0.0))
                for _, r in edited_exp.iterrows()
                if str(r.get("Description") or "").strip() or float(r.get("Amount (₦)") or 0.0) != 0.0
            ]
        with ec2:
            st.markdown("**Bank Lodgment (Optional)**")
            st_state.bank_name = st.text_input(
                "Bank Name", value=st_state.bank_name, key=f"bank_n_{st_state.file_name}"
            )
            st_state.bank_amount_deposited = st.number_input(
                "Amount Deposited (₦)",
                value=float(st_state.bank_amount_deposited),
                step=1000.0,
                format="%.2f",
                key=f"bank_a_{st_state.file_name}",
            )
            st_state.bank_payment_status = st.text_input(
                "Payment Status", value=st_state.bank_payment_status, key=f"bank_s_{st_state.file_name}"
            )

    # ---------------- SAVE & NAVIGATION BAR (Require Save Before Next!)
    st.divider()
    dirty_now = _has_unsaved_changes(cur_fn)
    if st_state.validation_errors:
        st.error(
            f"❌ **`{st_state.station_name}` failed the report integrity check and was NOT added to the download queue.**\n\n"
            + "\n".join(f"- {e}" for e in st_state.validation_errors)
            + "\n\nFix the figures above and click **💾 Save Station Report** again."
        )
    if dirty_now:
        st.warning(
            f"⚠️ **You have unsaved changes on `{st_state.station_name}`.** "
            "Please click **💾 Save Station Report** before clicking **Next** (or click **💾 Save & Next Remaining ➡️**)."
        )
    elif st_state.is_saved:
        st.success(
            f"✅ **`{st_state.station_name}` is saved, validated, and queued for Download All** (saved at {st_state.saved_at})."
        )
    if st_state.validation_warnings and not st_state.validation_errors:
        with st.expander(f"ℹ️ {len(st_state.validation_warnings)} validation note(s)"):
            for w in st_state.validation_warnings:
                st.caption(f"- {w}")

    act1, act2, act3, act4, act5 = st.columns([1.2, 2.0, 2.2, 1.4, 1.8])

    if act1.button("⬅️ Prev", use_container_width=True, disabled=(cur_idx == 0 or dirty_now)):
        st.session_state["selected_file"] = file_keys[cur_idx - 1]
        st.rerun()

    if act2.button("💾 Save Station Report", type="primary", use_container_width=True):
        with st.spinner(f"Generating & validating {st_state.station_name}..."):
            res = _save_station(cur_fn)
        if res.ok:
            st.toast(f"✅ Saved & queued {st_state.station_name}!")
        else:
            st.toast(f"❌ {st_state.station_name} failed validation — not queued.")
        st.rerun()

    remaining_after = [
        fn for fn in (file_keys[cur_idx + 1:] + file_keys[:cur_idx])
        if not states[fn].is_saved and fn != cur_fn
    ]
    if act3.button(
        "💾 Save & Next Remaining ➡️",
        use_container_width=True,
        disabled=(len(remaining_after) == 0 and st_state.is_saved and not dirty_now),
    ):
        with st.spinner(f"Generating & validating {st_state.station_name}..."):
            res = _save_station(cur_fn)
        if res.ok:
            st.toast(f"✅ Saved & queued {st_state.station_name}!")
            if remaining_after:
                st.session_state["selected_file"] = remaining_after[0]
        else:
            # Stay on this station so the user can see and fix the problem
            st.toast(f"❌ {st_state.station_name} failed validation — staying here to fix it.")
        st.rerun()

    if act4.button(
        "Next ➡️",
        use_container_width=True,
        disabled=(cur_idx >= len(file_keys) - 1 or dirty_now or not st_state.is_saved),
        help="Save the current station report first to enable Next.",
    ):
        st.session_state["selected_file"] = file_keys[cur_idx + 1]
        st.rerun()

    if _is_queued(cur_fn):
        act5.download_button(
            "⬇️ Download This .xlsx",
            data=_make_file_callable(_store(), cur_fn),
            file_name=cur_fn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click="ignore",
            key=f"dl_one_{cur_fn}",
        )


# ================================================================ TAB 2: BATCH GRID & EXCEL IMPORT
with tab_batch:
    st.subheader("Centralized Batch Entry & Excel Template Import/Export")
    bcol1, bcol2 = st.columns(2)
    with bcol1:
        st.markdown("##### 📥 Download Pre-Filled Master Batch Template (.xlsx)")
        st.caption(
            "Contains all loaded stations with their rolled-over Opening Meters, Prices, Opening Tank Dips, "
            "and POS Channels in a single workbook."
        )
        template_bytes = export_batch_template(states)
        st.download_button(
            "⬇️ Download Master Batch Entry Template (.xlsx)",
            data=template_bytes,
            file_name=f"SGR_Master_Batch_Template_{next(iter(states.values())).report_date.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with bcol2:
        st.markdown("##### 📤 Upload Filled Master Batch Template (.xlsx)")
        batch_upload = st.file_uploader(
            "Upload populated template to update all stations at once",
            type=["xlsx"],
            key="batch_tpl_uploader",
        )
        if batch_upload is not None:
            if st.button("⚡ Apply Batch Template to All Stations", type="primary", use_container_width=True):
                cnt, msgs = import_batch_template(batch_upload.getvalue(), states)
                prog = st.progress(0.0, text="Generating & validating station workbooks...")
                failed = []
                total_b = len(states)
                for i, fn in enumerate(list(states.keys())):
                    res = _save_station(fn)  # one workbook at a time; memory freed after each
                    if not res.ok:
                        failed.append(states[fn].station_name)
                    for t_idx in range(1, 5):
                        _bump_rev(f"sales_{fn}_{t_idx}")
                    _bump_rev(f"tanks_{fn}")
                    prog.progress((i + 1) / max(1, total_b), text=f"Validated {i + 1}/{total_b}: {fn}")
                prog.empty()
                out_msgs = [f"✅ {m}" for m in msgs]
                if failed:
                    out_msgs.append(
                        f"❌ {len(failed)} station(s) failed validation and were NOT queued: " + ", ".join(failed)
                    )
                st.session_state["upload_msgs"] = out_msgs
                st.rerun()


# ================================================================ TAB 3: SUMMARY & ZIP DOWNLOAD
with tab_export:
    st.subheader("Executive Summary & Download All")
    st.caption(
        "Every time you click **💾 Save Station Report**, the workbook is generated, checked for errors, and "
        "added to the download queue on the server's disk. When you're done, click **Download All** once."
    )

    def _queue_status(fn: str) -> str:
        s = states[fn]
        if s.validation_errors:
            return "❌ Failed check"
        if s.is_saved and _has_unsaved_changes(fn):
            return "🟡 Edited since save"
        if _is_queued(fn):
            return "📦 Queued"
        return "⏳ Not saved"

    summary_rows = _summary_rows(states, list(states.keys()))
    for row, fn in zip(summary_rows, states.keys()):
        row["Queue"] = _queue_status(fn)
    sum_df = pd.DataFrame(summary_rows)
    if not sum_df.empty:
        sum_df = sum_df[["Queue"] + [c for c in sum_df.columns if c != "Queue"]]
    st.dataframe(sum_df, use_container_width=True, hide_index=True, height=420)

    stale_fns = [fn for fn, s in states.items() if s.is_saved and _has_unsaved_changes(fn)]
    failed_fns = [fn for fn, s in states.items() if s.validation_errors]

    dl_c1, dl_c2 = st.columns([3, 2])
    with dl_c1:
        _render_download_all(states, key="dl_all_export")
    with dl_c2:
        if st.button(
            f"💾 Re-save {len(stale_fns)} Edited Station(s)",
            use_container_width=True,
            disabled=not stale_fns,
            help="Re-generates and re-validates stations that were changed after their last save.",
        ):
            prog = st.progress(0.0, text="Re-saving edited stations...")
            bad = []
            for i, fn in enumerate(stale_fns):
                if not _save_station(fn).ok:
                    bad.append(states[fn].station_name)
                prog.progress((i + 1) / len(stale_fns), text=f"Re-saved {i + 1}/{len(stale_fns)}: {fn}")
            prog.empty()
            if bad:
                st.session_state["upload_msgs"] = [f"❌ Failed validation (not queued): {', '.join(bad)}"]
            st.rerun()

    if failed_fns:
        with st.expander(f"❌ {len(failed_fns)} station(s) failed the integrity check", expanded=True):
            for fn in failed_fns:
                st.markdown(f"**{states[fn].station_name}**")
                for e in states[fn].validation_errors:
                    st.caption(f"- {e}")
