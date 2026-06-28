from __future__ import annotations

import json
from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

from config.settings import (
    DATA_DIR,
    LATEST_META_JSON,
    LATEST_REPORT_XLSX,
    LATEST_RESULT_CSV,
    LATEST_STATUS_DETAIL_CSV,
    LATEST_STATUS_SUMMARY_CSV,
)


def save_latest_dashboard(
    result_df: pd.DataFrame,
    report: BytesIO,
    total_docs: int,
    open_docs: int,
    action_counts: dict[str, int],
    status_summary_df: pd.DataFrame,
    status_detail_df: pd.DataFrame,
) -> str:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    result_df.to_csv(LATEST_RESULT_CSV, index=False, encoding="utf-8-sig")

    if status_summary_df is not None and not status_summary_df.empty:
        status_summary_df.to_csv(LATEST_STATUS_SUMMARY_CSV, index=False, encoding="utf-8-sig")

    if status_detail_df is not None and not status_detail_df.empty:
        status_detail_df.to_csv(LATEST_STATUS_DETAIL_CSV, index=False, encoding="utf-8-sig")

    with open(LATEST_REPORT_XLSX, "wb") as file:
        file.write(report.getvalue())

    last_updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

    meta = {
        "total_docs": int(total_docs),
        "open_docs": int(open_docs),
        "action_counts": action_counts,
        "last_updated": last_updated,
    }

    with open(LATEST_META_JSON, "w", encoding="utf-8") as file:
        json.dump(meta, file, ensure_ascii=False, indent=2)

    return last_updated


def load_latest_dashboard() -> tuple[pd.DataFrame | None, dict[str, Any] | None, bytes | None, pd.DataFrame | None, pd.DataFrame | None]:
    if not LATEST_RESULT_CSV.exists() or not LATEST_META_JSON.exists():
        return None, None, None, None, None

    result_df = pd.read_csv(LATEST_RESULT_CSV)

    with open(LATEST_META_JSON, "r", encoding="utf-8") as file:
        meta = json.load(file)

    report_bytes = LATEST_REPORT_XLSX.read_bytes() if LATEST_REPORT_XLSX.exists() else None

    status_summary_df = pd.read_csv(LATEST_STATUS_SUMMARY_CSV) if LATEST_STATUS_SUMMARY_CSV.exists() else None
    status_detail_df = pd.read_csv(LATEST_STATUS_DETAIL_CSV) if LATEST_STATUS_DETAIL_CSV.exists() else None

    return result_df, meta, report_bytes, status_summary_df, status_detail_df


def hydrate_session_from_latest() -> None:
    if st.session_state.result_df is not None:
        return

    result_df, meta, report_bytes, status_summary_df, status_detail_df = load_latest_dashboard()

    if result_df is None or meta is None:
        return

    st.session_state.result_df = result_df
    st.session_state.report = report_bytes
    st.session_state.total_docs = int(meta.get("total_docs", 0))
    st.session_state.open_docs = int(meta.get("open_docs", 0))
    st.session_state.action_counts = meta.get("action_counts", {})
    st.session_state.last_updated = meta.get("last_updated", "")
    st.session_state.status_summary_df = status_summary_df
    st.session_state.status_detail_df = status_detail_df
