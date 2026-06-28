from __future__ import annotations

import streamlit as st


def render_kpi_cards(action_counts: dict[str, int]) -> None:
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
