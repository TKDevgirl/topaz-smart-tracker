from __future__ import annotations

import streamlit as st

from services.trend_service import build_key_action_trends, build_trend_dashboard_data


def render_trend_dashboard() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">📈 Trend Dashboard</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-subtitle">Trend analytics is based on SQLite run history.</div>',
        unsafe_allow_html=True,
    )

    trend_df, action_trend_df = build_trend_dashboard_data()

    if trend_df is None or trend_df.empty or len(trend_df) < 2:
        run_count = 0 if trend_df is None or trend_df.empty else len(trend_df)

        st.markdown(
            f"""
            <div class="empty-state">
                <div class="empty-state-title">Trend will be available after 2 generations</div>
                Current historical runs: <b>{run_count}</b><br>
                Generate the dashboard again with the next updated Excel file to start trend comparison.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("</div>", unsafe_allow_html=True)
        return

    chart_df = trend_df.set_index("Run")[["Approval Rate", "Health Score", "Open Documents"]]
    st.line_chart(chart_df)

    key_action_df = build_key_action_trends(action_trend_df)

    if key_action_df is not None and not key_action_df.empty and len(key_action_df) >= 2:
        st.markdown("#### Key Action Trend")
        action_cols = [col for col in key_action_df.columns if col not in ["run_id", "Run"]]
        st.line_chart(key_action_df.set_index("Run")[action_cols])

    with st.expander("View trend data"):
        st.dataframe(trend_df, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)
