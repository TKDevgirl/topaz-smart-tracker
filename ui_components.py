import os

import pandas as pd
import streamlit as st

from config import APP_VERSION, LOGO_PATH
from utils import dataframe_to_excel_bytes, image_to_base64


def render_logo_html(size_class: str = "hero-logo") -> str:
    logo64 = image_to_base64(LOGO_PATH)

    if logo64:
        return f'<img class="{size_class}" src="data:image/png;base64,{logo64}">'

    return '<div style="font-size:62px;">💎</div>'


def init_session_state():
    defaults = {
        "logged_in": False,
        "role": "viewer",
        "username": "",
        "result_df": None,
        "report": None,
        "total_docs": 0,
        "open_docs": 0,
        "action_counts": {},
        "last_updated": "",
        "status_summary_df": None,
        "status_detail_df": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_sidebar():
    with st.sidebar:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=150)
        else:
            st.markdown("## 💎")

        st.markdown('<div class="sidebar-logo-title">TOPAZ</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sidebar-subtitle">Smart Document Tracker {APP_VERSION}</div>', unsafe_allow_html=True)

        st.divider()

        if not st.session_state.logged_in:
            username = st.text_input("Username")

            if st.button("Login", use_container_width=True):
                st.session_state.username = username.strip()

                if username.strip().lower() in ["pavinee", "admin"]:
                    st.session_state.role = "admin"
                else:
                    st.session_state.role = "viewer"

                st.session_state.logged_in = True
                st.rerun()

        else:
            st.markdown(
                f"""
                <div class="sidebar-card">
                    <b>🔑 Role</b><br>{st.session_state.role.title()}
                </div>
                """,
                unsafe_allow_html=True,
            )

            if st.button("Logout", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.role = "viewer"
                st.session_state.username = ""
                st.rerun()

        st.divider()
        st.markdown('<div class="nav-active">🏠 Dashboard</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">⬇ Download Report</div>', unsafe_allow_html=True)
        st.divider()
        st.caption("Shared Dashboard")


def render_header():
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-grid">
                <div>{render_logo_html()}</div>
                <div>
                    <div class="hero-title">Topaz Smart Document Tracker</div>
                    <div class="hero-subtitle">TOPAZ BKK1 | ICT Document Control Dashboard</div>
                </div>
                <div class="hero-badge">Shared Dashboard</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_cards(action_counts):
    c1, c2, c3, c4, c5 = st.columns(5)

    cards = [
        ("📁", "Total Documents", st.session_state.total_docs, "All documents", "#7c3aed"),
        ("🕒", "Open / On Progress", st.session_state.open_docs, "Documents", "#2563eb"),
        ("▶️", "Open & On Process", action_counts.get("OPEN & ON PROCESS", 0), "Waiting review", "#16a34a"),
        ("🔄", "Need Update", action_counts.get("UPDATE TRACKING TO CLOSED", 0), "Update tracking", "#f97316"),
        ("⚠️", "Overdue", action_counts.get("OVERDUE / FOLLOW UP", 0), "Follow up", "#ef4444"),
    ]

    for col, (icon, title, value, sub, color) in zip([c1, c2, c3, c4, c5], cards):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card" style="--accent:{color};">
                    <div class="kpi-icon">{icon}</div>
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_action_summary(action_counts):
    chart_col, table_col = st.columns([1.4, 1])

    with chart_col:
        st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)

        summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])

        if not summary_df.empty:
            st.bar_chart(summary_df.set_index("Action"))
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        else:
            st.info("No action data found.")

        st.markdown("</div>", unsafe_allow_html=True)

    with table_col:
        st.markdown('<div class="panel"><div class="panel-title">🧭 Quick Action</div>', unsafe_allow_html=True)

        quick_items = [
            ("Returned by NV5", action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)),
            ("Need update closed", action_counts.get("UPDATE TRACKING TO CLOSED", 0)),
            ("Overdue follow up", action_counts.get("OVERDUE / FOLLOW UP", 0)),
            ("Not found in Takenaka", action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)),
        ]

        for label, count in quick_items:
            st.markdown(
                f'<div class="quick-row"><span>{label}</span><span class="count-pill">{count}</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)


def render_status_summary_panel():
    status_summary_df = st.session_state.get("status_summary_df", None)

    st.markdown(
        '<div class="panel"><div class="panel-title">📌 RFA / RFI Status Summary by Category (Latest Revision)</div>',
        unsafe_allow_html=True,
    )

    if status_summary_df is not None and not status_summary_df.empty:
        def highlight_status(row):
            status = str(row.get("Status", ""))

            if status == "Open":
                return ["background-color: #dbeafe; color: #1e3a8a; font-weight: 800"] * len(row)
            if status == "On Progress":
                return ["background-color: #fef3c7; color: #92400e; font-weight: 800"] * len(row)
            if status == "Approved":
                return ["background-color: #dcfce7; color: #166534; font-weight: 800"] * len(row)
            if status == "Total Document":
                return ["background-color: #ede9fe; color: #4c1d95; font-weight: 950"] * len(row)

            return [""] * len(row)

        rfa_summary_df = status_summary_df[status_summary_df["Document Type"] == "RFA"].drop(columns=["Document Type"], errors="ignore")
        rfi_summary_df = status_summary_df[status_summary_df["Document Type"] == "RFI"].drop(columns=["Document Type"], errors="ignore")

        st.markdown("#### 📋 RFA Status Summary")
        st.dataframe(
            rfa_summary_df.style.apply(highlight_status, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.write("")

        st.markdown("#### 📄 RFI Status Summary")
        if not rfi_summary_df.empty:
            st.dataframe(
                rfi_summary_df.style.apply(highlight_status, axis=1),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No RFI summary data found.")

        st.download_button(
            label="⬇️ Export RFA / RFI Status Summary",
            data=dataframe_to_excel_bytes(status_summary_df, "Status Summary"),
            file_name="Topaz_RFA_RFI_Status_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )

    else:
        st.info("No status summary available. Please generate dashboard again with Tracking_document.xlsx.")

    st.markdown("</div>", unsafe_allow_html=True)


def render_document_action_list(df):
    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)

    filter_col, search_col, download_col = st.columns([1, 2, 0.8])

    with filter_col:
        selected_action = st.selectbox(
            "Filter by Action",
            ["All"] + sorted(df["Action"].dropna().unique().tolist()),
        )

    with search_col:
        search = st.text_input("Search Document No / Document Name")

    filtered_df = df.copy()

    if selected_action != "All":
        filtered_df = filtered_df[filtered_df["Action"] == selected_action]

    if search:
        filtered_df = filtered_df[
            filtered_df.astype(str)
            .apply(lambda x: x.str.contains(search, case=False, na=False))
            .any(axis=1)
        ]

    st.dataframe(filtered_df, use_container_width=True, height=480)

    with download_col:
        st.write("")
        st.write("")

        if st.session_state.report is not None:
            st.download_button(
                label="⬇️ Download Excel Report",
                data=st.session_state.report,
                file_name="Open_On_Process_Compare.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)
