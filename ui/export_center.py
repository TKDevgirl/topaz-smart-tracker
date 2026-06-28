from __future__ import annotations

import streamlit as st

from core.utils import dataframe_to_excel_bytes
from repositories.dashboard_repository import list_dashboard_runs
from services.export_center_service import build_database_backup_excel
from services.weekly_report_service import build_weekly_report_excel


def render_export_center() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">📦 Export Center</div>',
        unsafe_allow_html=True,
    )

    result_df = st.session_state.get("result_df", None)
    status_summary_df = st.session_state.get("status_summary_df", None)
    status_detail_df = st.session_state.get("status_detail_df", None)
    action_counts = st.session_state.get("action_counts", {})

    runs_df = list_dashboard_runs()

    col1, col2, col3 = st.columns(3)

    with col1:
        if result_df is not None:
            st.download_button(
                "⬇️ Export Action List",
                data=dataframe_to_excel_bytes(result_df, "Action List"),
                file_name="Topaz_Action_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with col2:
        if status_summary_df is not None:
            st.download_button(
                "⬇️ Export Status Summary",
                data=dataframe_to_excel_bytes(status_summary_df, "Status Summary"),
                file_name="Topaz_Status_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with col3:
        weekly_bytes = build_weekly_report_excel(result_df, status_summary_df, action_counts)
        st.download_button(
            "⬇️ Export Weekly Report",
            data=weekly_bytes,
            file_name="Topaz_Weekly_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    backup_bytes = build_database_backup_excel(
        runs_df=runs_df,
        result_df=result_df,
        status_summary_df=status_summary_df,
        status_detail_df=status_detail_df,
    )

    st.download_button(
        "🗄️ Export SQLite Data Backup",
        data=backup_bytes,
        file_name="Topaz_SQLite_Backup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    st.markdown("</div>", unsafe_allow_html=True)
