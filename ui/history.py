from __future__ import annotations

import streamlit as st

from services.history_service import list_history_snapshots


def render_history_panel() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">🕘 Dashboard History</div>',
        unsafe_allow_html=True,
    )

    history_df = list_history_snapshots()

    if history_df.empty:
        st.info("No history snapshot yet. Generate dashboard to create the first snapshot.")
    else:
        st.dataframe(history_df, use_container_width=True, hide_index=True, height=220)

    st.markdown("</div>", unsafe_allow_html=True)
