from __future__ import annotations

import streamlit as st

from services.analytics_service import calculate_category_totals, calculate_project_health
from services.insight_service import generate_executive_insights, generate_recommended_actions


def render_executive_dashboard() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)
    action_counts = st.session_state.get("action_counts", {})

    health = calculate_project_health(status_summary_df, action_counts)

    st.markdown(
        '<div class="panel"><div class="panel-title">🚀 Executive AI Overview</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Project Health", f"{health['emoji']} {health['level']}", f"{health['score']}/100")
    c2.metric("Approval Rate", f"{health['approved_rate']}%")
    c3.metric("Open", health["open"])
    c4.metric("Overdue", health["overdue"])

    st.markdown("#### 🤖 AI-style Insights")
    for insight in generate_executive_insights(status_summary_df, action_counts):
        st.markdown(f"- {insight}")

    st.markdown("#### ✅ Recommended Actions")
    for action in generate_recommended_actions(status_summary_df, action_counts):
        st.markdown(f"- {action}")

    st.markdown("</div>", unsafe_allow_html=True)


def render_category_analytics() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)
    category_df = calculate_category_totals(status_summary_df)

    st.markdown(
        '<div class="panel"><div class="panel-title">📈 Category Analytics</div>',
        unsafe_allow_html=True,
    )

    if category_df is not None and not category_df.empty:
        st.bar_chart(category_df.set_index("Category"))
        st.dataframe(category_df, use_container_width=True, hide_index=True)
    else:
        st.info("No category analytics available yet.")

    st.markdown("</div>", unsafe_allow_html=True)
