from __future__ import annotations

import streamlit as st

from core.utils import dataframe_to_excel_bytes
from services.search_service import build_document_timeline, get_filter_options, search_documents


def render_smart_search_panel() -> None:
    detail_df = st.session_state.get("status_detail_df", None)

    st.markdown(
        '<div class="panel"><div class="panel-title">🔍 Smart Document Search</div>',
        unsafe_allow_html=True,
    )

    if detail_df is None or detail_df.empty:
        st.info("No document detail data available yet. Please generate dashboard first.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    options = get_filter_options(detail_df)

    col_1, col_2, col_3, col_4 = st.columns([2, 1, 1, 1])

    with col_1:
        query = st.text_input(
            "Search Document No / Name / Info",
            placeholder="Example: MAT, RFI, 00023, PDU, Fiber",
            key="smart_search_query",
        )

    with col_2:
        document_type = st.selectbox(
            "Type",
            options["document_types"],
            key="smart_search_type",
        )

    with col_3:
        category = st.selectbox(
            "Category",
            options["categories"],
            key="smart_search_category",
        )

    with col_4:
        status = st.selectbox(
            "Status",
            options["statuses"],
            key="smart_search_status",
        )

    result_df = search_documents(
        detail_df,
        query=query,
        document_type=document_type,
        category=category,
        status=status,
    )

    st.caption(f"Search result: {len(result_df)} document(s)")
    st.dataframe(result_df, use_container_width=True, hide_index=True, height=330)

    st.download_button(
        label="⬇️ Export Search Result",
        data=dataframe_to_excel_bytes(result_df, "Search Result"),
        file_name="Topaz_Smart_Search_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )

    st.markdown("</div>", unsafe_allow_html=True)


def render_document_timeline_panel() -> None:
    detail_df = st.session_state.get("status_detail_df", None)

    st.markdown(
        '<div class="panel"><div class="panel-title">🧾 Document Timeline</div>',
        unsafe_allow_html=True,
    )

    if detail_df is None or detail_df.empty or "Document No" not in detail_df.columns:
        st.info("No document detail data available yet.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    document_list = sorted(detail_df["Document No"].dropna().astype(str).unique().tolist())

    if not document_list:
        st.info("No document number found.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    selected_doc = st.selectbox(
        "Select Document No",
        document_list,
        key="timeline_document_no",
    )

    timeline_df = build_document_timeline(detail_df, selected_doc)

    if timeline_df.empty:
        st.info("No timeline data found for this document.")
    else:
        st.dataframe(timeline_df, use_container_width=True, hide_index=True)

        # Simple readable timeline cards
        for _, row in timeline_df.iterrows():
            st.markdown(
                f"""
                <div class="quick-row">
                    <span><b>{row.get('Document No', '')}</b> | Rev {row.get('Revision', '')} | {row.get('Category', '')}</span>
                    <span class="count-pill">{row.get('Status', '')}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)
