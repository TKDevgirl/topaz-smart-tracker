from __future__ import annotations

import streamlit as st

from services.weekly_report_service import build_weekly_report_excel


def render_weekly_report_panel() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">📅 Weekly Report</div>',
        unsafe_allow_html=True,
    )

    result_df = st.session_state.get("result_df", None)
    status_summary_df = st.session_state.get("status_summary_df", None)
    action_counts = st.session_state.get("action_counts", {})

    if result_df is None or status_summary_df is None:
        st.info("No dashboard data available yet. Generate dashboard first.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    report_bytes = build_weekly_report_excel(result_df, status_summary_df, action_counts)

    st.markdown(
        "Generate an Excel weekly report for PM meeting including KPI, insights, actions, status summary, trend, and action list."
    )

    st.download_button(
        label="⬇️ Download Weekly Report Excel",
        data=report_bytes,
        file_name="Topaz_Weekly_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    st.markdown("</div>", unsafe_allow_html=True)
