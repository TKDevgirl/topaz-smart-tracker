import json
from datetime import datetime

import pandas as pd
import streamlit as st

from config import (
    DATA_DIR,
    LATEST_CSV,
    LATEST_META,
    LATEST_REPORT,
    LATEST_STATUS_DETAIL,
    LATEST_STATUS_SUMMARY,
)


def save_latest_dashboard(
    df: pd.DataFrame,
    report,
    total_docs: int,
    open_docs: int,
    action_counts: dict,
    status_summary_df: pd.DataFrame,
    status_detail_df: pd.DataFrame,
) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    df.to_csv(LATEST_CSV, index=False, encoding="utf-8-sig")

    if status_summary_df is not None and not status_summary_df.empty:
        status_summary_df.to_csv(LATEST_STATUS_SUMMARY, index=False, encoding="utf-8-sig")

    if status_detail_df is not None and not status_detail_df.empty:
        status_detail_df.to_csv(LATEST_STATUS_DETAIL, index=False, encoding="utf-8-sig")

    with open(LATEST_REPORT, "wb") as f:
        f.write(report.getvalue())

    meta = {
        "total_docs": int(total_docs),
        "open_docs": int(open_docs),
        "action_counts": action_counts,
        "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
    }

    with open(LATEST_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def load_latest_dashboard():
    if not LATEST_CSV.exists() or not LATEST_META.exists():
        return None, None, None, None, None

    df = pd.read_csv(LATEST_CSV)

    with open(LATEST_META, "r", encoding="utf-8") as f:
        meta = json.load(f)

    report_bytes = LATEST_REPORT.read_bytes() if LATEST_REPORT.exists() else None

    status_summary_df = None
    if LATEST_STATUS_SUMMARY.exists():
        status_summary_df = pd.read_csv(LATEST_STATUS_SUMMARY)

    status_detail_df = None
    if LATEST_STATUS_DETAIL.exists():
        status_detail_df = pd.read_csv(LATEST_STATUS_DETAIL)

    return df, meta, report_bytes, status_summary_df, status_detail_df


def hydrate_session_from_latest() -> None:
    if st.session_state.result_df is not None:
        return

    df, meta, report_bytes, status_summary_df, status_detail_df = load_latest_dashboard()

    if df is None or meta is None:
        return

    st.session_state.result_df = df
    st.session_state.report = report_bytes
    st.session_state.total_docs = int(meta.get("total_docs", 0))
    st.session_state.open_docs = int(meta.get("open_docs", 0))
    st.session_state.action_counts = meta.get("action_counts", {})
    st.session_state.last_updated = meta.get("last_updated", "")
    st.session_state.status_summary_df = status_summary_df
    st.session_state.status_detail_df = status_detail_df
