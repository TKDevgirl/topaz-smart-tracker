from __future__ import annotations

import streamlit as st

from core.storage import get_dashboard_runs, load_run_to_session


def render_database_panel() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">🗄️ SQLite Database History</div>',
        unsafe_allow_html=True,
    )

    runs_df = get_dashboard_runs()

    if runs_df.empty:
        st.info("No database history yet. Generate dashboard once to create the first SQLite run.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    st.dataframe(runs_df, use_container_width=True, hide_index=True, height=260)

    run_ids = runs_df["Run ID"].tolist()

    col_1, col_2 = st.columns([1, 3])

    with col_1:
        selected_run_id = st.selectbox(
            "Load Run ID",
            run_ids,
            key="sqlite_run_id",
        )

    with col_2:
        st.write("")
        st.write("")
        if st.button("Load Selected Snapshot", use_container_width=False):
            load_run_to_session(int(selected_run_id))
            st.rerun()

    st.caption("SQLite stores each Generate as a new run_id. The dashboard loads the latest run by default.")

    st.markdown("</div>", unsafe_allow_html=True)
