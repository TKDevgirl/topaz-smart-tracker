from __future__ import annotations

import pandas as pd
import streamlit as st

from services.analytics_service import calculate_category_totals, calculate_project_health
from services.insight_service import generate_executive_insights, generate_recommended_actions


def _progress_bar(value: float, max_value: float = 100.0) -> str:
    value = max(0.0, min(float(value), max_value))
    width = 0 if max_value == 0 else round((value / max_value) * 100, 1)

    return f"""
    <div style="width:100%; background:#e5e7eb; border-radius:999px; height:13px; overflow:hidden;">
        <div style="width:{width}%; background:linear-gradient(90deg,#ef4444,#f97316,#22c55e); height:13px;"></div>
    </div>
    """


def _metric_box(title: str, value: str, subtitle: str = "") -> str:
    return f"""
    <div style="
        background:#ffffff;
        border:1px solid #e2e8f0;
        border-radius:18px;
        padding:18px;
        box-shadow:0 8px 20px rgba(15,23,42,.06);
        min-height:115px;">
        <div style="font-size:13px;color:#64748b;font-weight:800;">{title}</div>
        <div style="font-size:34px;font-weight:950;color:#0f172a;line-height:1.15;margin-top:8px;">{value}</div>
        <div style="font-size:12px;color:#64748b;margin-top:4px;">{subtitle}</div>
    </div>
    """


def _top_items(items: list[str], limit: int = 4) -> list[str]:
    return items[:limit]


def render_executive_dashboard() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)
    action_counts = st.session_state.get("action_counts", {})

    health = calculate_project_health(status_summary_df, action_counts)
    insights = _top_items(generate_executive_insights(status_summary_df, action_counts), 4)
    actions = _top_items(generate_recommended_actions(status_summary_df, action_counts), 3)

    st.markdown(
        '<div class="panel"><div class="panel-title">🚀 Executive Summary</div>',
        unsafe_allow_html=True,
    )

    status_color = "#ef4444" if health["level"] == "Risk" else "#f59e0b" if health["level"] == "Watch" else "#22c55e"

    top_left, top_right = st.columns([1.05, 1])

    with top_left:
        st.markdown(
            f"""
            <div style="
                background:linear-gradient(135deg,#0f172a,#1e3a8a);
                color:white;
                border-radius:24px;
                padding:26px;
                box-shadow:0 16px 35px rgba(15,23,42,.18);">
                <div style="font-size:14px;color:#bfdbfe;font-weight:800;">Project Health</div>
                <div style="display:flex;align-items:center;gap:14px;margin-top:10px;">
                    <div style="font-size:50px;">{health['emoji']}</div>
                    <div>
                        <div style="font-size:38px;font-weight:950;">{health['level']}</div>
                        <div style="font-size:15px;color:#dbeafe;">Score {health['score']}/100</div>
                    </div>
                </div>
                <div style="margin-top:18px;">{_progress_bar(float(health['score']))}</div>
                <div style="font-size:13px;color:#dbeafe;margin-top:10px;">
                    Approval Rate {health['approved_rate']}% | Open {health['open']} | Overdue {health['overdue']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with top_right:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(_metric_box("Approved", str(health["approved"]), f"{health['approved_rate']}% approval"), unsafe_allow_html=True)
        with c2:
            st.markdown(_metric_box("Open", str(health["open"]), "Require monitoring"), unsafe_allow_html=True)

        st.write("")

        c3, c4 = st.columns(2)
        with c3:
            st.markdown(_metric_box("Returned", str(health["returned"]), "Need resubmission"), unsafe_allow_html=True)
        with c4:
            st.markdown(_metric_box("Overdue", str(health["overdue"]), "Urgent follow-up"), unsafe_allow_html=True)

    st.write("")

    insight_col, action_col = st.columns(2)

    with insight_col:
        st.markdown("#### 🤖 Key Findings")
        for insight in insights:
            st.markdown(f"- {insight}")

    with action_col:
        st.markdown("#### ✅ Priority Actions")
        for idx, action in enumerate(actions, start=1):
            st.markdown(f"**{idx}.** {action}")

    st.markdown("</div>", unsafe_allow_html=True)


def render_category_analytics() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)
    category_df = calculate_category_totals(status_summary_df)

    st.markdown(
        '<div class="panel"><div class="panel-title">📈 Category Analytics</div>',
        unsafe_allow_html=True,
    )

    if category_df is not None and not category_df.empty:
        category_df = category_df.sort_values("Count", ascending=False)
        st.bar_chart(category_df.set_index("Category"))

        top_category = category_df.iloc[0]
        st.caption(
            f"Highest volume: {top_category['Category']} with {int(top_category['Count'])} document(s)."
        )
    else:
        st.info("No category analytics available yet.")

    st.markdown("</div>", unsafe_allow_html=True)
