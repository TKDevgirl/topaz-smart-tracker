from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from config.settings import APP_ICON, APP_TITLE
from core.database import init_database
from core.session import init_session_state
from core.storage import hydrate_session_from_latest, save_latest_dashboard
from services.report_service import generate_report
from services.history_service import save_history_snapshot
from ui.alerts import render_alert_panel
from ui.enterprise import render_alert_center, render_run_timeline_panel, render_snapshot_compare_panel, render_sync_bar
from ui.export_center import render_export_center
from ui.cards import render_kpi_cards
from ui.layout import render_header, render_role_message
from ui.sidebar import render_sidebar
from ui.executive import render_category_analytics, render_executive_dashboard
from ui.history import render_history_panel
from ui.database import render_database_panel
from ui.search import render_document_timeline_panel, render_smart_search_panel
from ui.trend import render_trend_dashboard
from ui.weekly_report import render_weekly_report_panel
from ui.tables import (
    render_action_summary,
    render_document_action_list,
    render_status_summary_panel,
)
from ui.theme import apply_theme


def render_admin_upload() -> None:
    render_role_message()

    if st.session_state.role != "admin":
        return

    col_upload_1, col_upload_2, col_button = st.columns([1.15, 1.15, 0.65])

    with col_upload_1:
        tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])

    with col_upload_2:
        takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

    with col_button:
        st.write("")
        st.write("")
        generate_clicked = st.button("🚀 Generate", type="primary", use_container_width=True)

    if tracking_file and takenaka_file and generate_clicked:
        with st.spinner("Reading files and generating dashboard..."):
            report, total_docs, open_docs, rows, status_summary_df, status_detail_df = generate_report(
                tracking_file,
                takenaka_file,
            )

        result_df = pd.DataFrame(rows)
        action_counts = result_df["Action"].value_counts().to_dict() if not result_df.empty else {}

        last_updated = save_latest_dashboard(
            result_df,
            report,
            total_docs,
            open_docs,
            action_counts,
            status_summary_df,
            status_detail_df,
        )

        st.session_state.result_df = result_df
        st.session_state.report = report
        st.session_state.total_docs = total_docs
        st.session_state.open_docs = open_docs
        st.session_state.action_counts = action_counts
        st.session_state.status_summary_df = status_summary_df
        st.session_state.status_detail_df = status_detail_df
        st.session_state.last_updated = last_updated or datetime.now().strftime("%d-%b-%Y %H:%M:%S")

        save_history_snapshot(status_summary_df, result_df)

        st.rerun()


def render_dashboard() -> None:
    if st.session_state.result_df is None:
        st.markdown(
            '<div class="warning-box">⚠️ No shared dashboard data available yet. Please login as Admin and generate dashboard once.</div>',
            unsafe_allow_html=True,
        )
        return

    result_df = st.session_state.result_df
    action_counts = st.session_state.action_counts
    last_updated = st.session_state.get("last_updated", "")

    if last_updated:
        st.success(f"Dashboard loaded successfully ✅ | Last updated: {last_updated}")
    else:
        st.success("Dashboard loaded successfully ✅")

    render_kpi_cards(action_counts)

    st.write("")
    st.write("")

    render_alert_center()
    render_executive_dashboard()
    render_status_summary_panel()
    render_category_analytics()
    render_smart_search_panel()
    render_document_timeline_panel()
    render_action_summary(action_counts)
    render_document_action_list(result_df)
    render_trend_dashboard()
    render_snapshot_compare_panel()
    render_run_timeline_panel()
    render_database_panel()
    render_export_center()


def main() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon=APP_ICON,
        layout="wide",
    )

    init_database()
    apply_theme()
    init_session_state()
    render_sidebar()
    hydrate_session_from_latest()
    render_header()
    if st.session_state.result_df is not None:
        render_sync_bar()
    render_admin_upload()
    render_dashboard()


if __name__ == "__main__":
    main()
