from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

from config.settings import DATA_DIR, LATEST_REPORT_XLSX
from repositories.dashboard_repository import (
    list_dashboard_runs,
    load_dashboard_run,
    save_dashboard_run,
)
from services.analytics_service import calculate_project_health


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

    health = calculate_project_health(status_summary_df, action_counts)
    uploaded_by = st.session_state.get("username", "")

    save_dashboard_run(
        result_df=result_df,
        status_summary_df=status_summary_df,
        status_detail_df=status_detail_df,
        total_docs=total_docs,
        open_docs=open_docs,
        approval_rate=float(health.get("approved_rate", 0)),
        health_score=float(health.get("score", 0)),
        uploaded_by=uploaded_by,
    )

    # Keep the latest Excel report file for the download button.
    with open(LATEST_REPORT_XLSX, "wb") as file:
        file.write(report.getvalue())

    return datetime.now().strftime("%d-%b-%Y %H:%M:%S")


def hydrate_session_from_latest() -> None:
    if st.session_state.result_df is not None:
        return

    meta, result_df, status_summary_df, status_detail_df = load_dashboard_run()

    if meta is None:
        return

    action_counts = result_df["Action"].value_counts().to_dict() if result_df is not None and not result_df.empty else {}

    st.session_state.result_df = result_df
    st.session_state.report = LATEST_REPORT_XLSX.read_bytes() if LATEST_REPORT_XLSX.exists() else None
    st.session_state.total_docs = int(meta.get("total_docs", 0))
    st.session_state.open_docs = int(meta.get("open_docs", 0))
    st.session_state.action_counts = action_counts
    st.session_state.last_updated = meta.get("run_time", "")
    st.session_state.status_summary_df = status_summary_df
    st.session_state.status_detail_df = status_detail_df


def get_dashboard_runs() -> pd.DataFrame:
    return list_dashboard_runs()


def load_run_to_session(run_id: int) -> None:
    meta, result_df, status_summary_df, status_detail_df = load_dashboard_run(run_id)

    if meta is None:
        return

    action_counts = result_df["Action"].value_counts().to_dict() if result_df is not None and not result_df.empty else {}

    st.session_state.result_df = result_df
    st.session_state.report = LATEST_REPORT_XLSX.read_bytes() if LATEST_REPORT_XLSX.exists() else None
    st.session_state.total_docs = int(meta.get("total_docs", 0))
    st.session_state.open_docs = int(meta.get("open_docs", 0))
    st.session_state.action_counts = action_counts
    st.session_state.last_updated = meta.get("run_time", "")
    st.session_state.status_summary_df = status_summary_df
    st.session_state.status_detail_df = status_detail_df
