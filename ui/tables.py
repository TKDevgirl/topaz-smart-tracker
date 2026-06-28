from __future__ import annotations

import pandas as pd
import streamlit as st

from core.utils import dataframe_to_excel_bytes


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


def render_action_summary(action_counts: dict[str, int]) -> None:
    st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)

    summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])

    if summary_df.empty:
        st.info("No action data found.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    summary_df = summary_df.sort_values("Count", ascending=False)
    max_count = max(summary_df["Count"].max(), 1)

    for _, row in summary_df.iterrows():
        action = row["Action"]
        count = int(row["Count"])
        width = round((count / max_count) * 100, 1)

        st.markdown(
            f"""
            <div style="display:grid;grid-template-columns:260px 1fr 48px;gap:14px;align-items:center;margin-bottom:12px;">
                <div style="font-weight:800;color:#334155;">{action}</div>
                <div style="height:14px;background:#e5e7eb;border-radius:999px;overflow:hidden;">
                    <div style="height:14px;width:{width}%;background:linear-gradient(90deg,#2563eb,#7c3aed);"></div>
                </div>
                <div class="count-pill" style="text-align:center;">{count}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">🧭 Quick Action</div>', unsafe_allow_html=True)

    quick_items = [
        ("Returned by NV5", action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)),
        ("Need update closed", action_counts.get("UPDATE TRACKING TO CLOSED", 0)),
        ("Overdue follow up", action_counts.get("OVERDUE / FOLLOW UP", 0)),
        ("Not found in Takenaka", action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)),
    ]

    q1, q2, q3, q4 = st.columns(4)
    for col, (label, count) in zip([q1, q2, q3, q4], quick_items):
        with col:
            st.markdown(
                f"""
                <div class="quick-row">
                    <span>{label}</span>
                    <span class="count-pill">{count}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)


def render_status_summary_panel() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)

    st.markdown(
        '<div class="panel"><div class="panel-title">📌 RFA / RFI Status Summary by Category (Latest Revision)</div>',
        unsafe_allow_html=True,
    )

    if status_summary_df is not None and not status_summary_df.empty:
        rfa_summary_df = status_summary_df[
            status_summary_df["Document Type"] == "RFA"
        ].drop(columns=["Document Type"], errors="ignore")

        rfi_summary_df = status_summary_df[
            status_summary_df["Document Type"] == "RFI"
        ].drop(columns=["Document Type"], errors="ignore")

        tab_rfa, tab_rfi, tab_all = st.tabs(["📋 RFA", "📄 RFI", "📊 Combined"])

        with tab_rfa:
            st.dataframe(
                rfa_summary_df.style.apply(highlight_status, axis=1),
                use_container_width=True,
                hide_index=True,
            )

        with tab_rfi:
            if not rfi_summary_df.empty:
                st.dataframe(
                    rfi_summary_df.style.apply(highlight_status, axis=1),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("No RFI summary data found.")

        with tab_all:
            st.dataframe(
                status_summary_df.style.apply(highlight_status, axis=1),
                use_container_width=True,
                hide_index=True,
            )

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


def render_document_action_list(df: pd.DataFrame) -> None:
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
