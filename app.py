import pandas as pd
import streamlit as st

from config import APP_TITLE
from report_generator import generate_report
from storage import hydrate_session_from_latest, save_latest_dashboard
from style import apply_style
from ui_components import (
    init_session_state,
    render_action_summary,
    render_document_action_list,
    render_header,
    render_kpi_cards,
    render_sidebar,
    render_status_summary_panel,
)


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="💎",
    layout="wide",
)

apply_style()
init_session_state()
render_sidebar()
hydrate_session_from_latest()
render_header()


# =========================================================
# ADMIN UPLOAD
# =========================================================
if st.session_state.role == "admin":
    st.markdown(
        '<div class="info-box">👩‍💼 Admin mode: Admin can upload files, generate dashboard, and update shared data.</div>',
        unsafe_allow_html=True,
    )

    col_upload_1, col_upload_2, col_button = st.columns([1, 1, 0.8])

    with col_upload_1:
        tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])

    with col_upload_2:
        takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

    with col_button:
        st.write("")
        st.write("")
        generate_clicked = st.button("🚀 Generate Dashboard", type="primary", use_container_width=True)

    if tracking_file and takenaka_file and generate_clicked:
        with st.spinner("Reading files and generating dashboard..."):
            report, total_docs, open_docs, rows, status_summary_df, status_detail_df = generate_report(
                tracking_file,
                takenaka_file,
            )

        df = pd.DataFrame(rows)
        action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}

        save_latest_dashboard(
            df,
            report,
            total_docs,
            open_docs,
            action_counts,
            status_summary_df,
            status_detail_df,
        )

        st.session_state.result_df = df
        st.session_state.report = report
        st.session_state.total_docs = total_docs
        st.session_state.open_docs = open_docs
        st.session_state.action_counts = action_counts
        st.session_state.status_summary_df = status_summary_df
        st.session_state.status_detail_df = status_detail_df

        from datetime import datetime
        st.session_state.last_updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

        st.rerun()

else:
    st.markdown(
        '<div class="viewer-box">ℹ️ Viewer mode: dashboard is read-only. Only Admin can upload files and update shared data.</div>',
        unsafe_allow_html=True,
    )


# =========================================================
# DASHBOARD VIEW
# =========================================================
if st.session_state.result_df is not None:
    df = st.session_state.result_df
    action_counts = st.session_state.action_counts
    last_updated = st.session_state.get("last_updated", "")

    if last_updated:
        st.success(f"Dashboard loaded successfully ✅ | Last updated: {last_updated}")
    else:
        st.success("Dashboard loaded successfully ✅")

    render_kpi_cards(action_counts)

    st.write("")
    st.write("")

    render_action_summary(action_counts)
    render_status_summary_panel()
    render_document_action_list(df)

else:
    st.markdown(
        '<div class="warning-box">⚠️ No shared dashboard data available yet. Please login as Admin and generate dashboard once.</div>',
        unsafe_allow_html=True,
    )
