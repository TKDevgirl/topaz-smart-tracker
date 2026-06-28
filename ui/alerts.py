from __future__ import annotations

import streamlit as st

from services.alert_service import build_alerts


def render_alert_panel() -> None:
    action_counts = st.session_state.get("action_counts", {})
    alerts = build_alerts(action_counts)

    st.markdown(
        '<div class="panel"><div class="panel-title">🔔 Alert Rules</div>',
        unsafe_allow_html=True,
    )

    for alert in alerts:
        level = alert["level"]
        title = alert["title"]
        message = alert["message"]

        if level == "critical":
            st.error(f"🚨 {title}: {message}")
        elif level == "warning":
            st.warning(f"⚠️ {title}: {message}")
        elif level == "info":
            st.info(f"ℹ️ {title}: {message}")
        else:
            st.success(f"✅ {title}: {message}")

    st.caption("Rules: Overdue > 0, Returned > 5, Need Update > 0")

    st.markdown("</div>", unsafe_allow_html=True)
