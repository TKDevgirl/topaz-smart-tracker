from __future__ import annotations

import os

import streamlit as st

from config.settings import APP_VERSION, LOGO_PATH
from core.session import login, logout


def render_sidebar() -> None:
    with st.sidebar:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=150)
        else:
            st.markdown("## 💎")

        st.markdown('<div class="sidebar-logo-title">TOPAZ</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="sidebar-subtitle">Smart Document Tracker {APP_VERSION}</div>',
            unsafe_allow_html=True,
        )

        st.divider()

        if not st.session_state.logged_in:
            username = st.text_input("Username")

            if st.button("Login", use_container_width=True):
                login(username)
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
                logout()
                st.rerun()

        st.divider()
        st.markdown('<div class="nav-active">🏠 Dashboard</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">⬇ Download Report</div>', unsafe_allow_html=True)
        st.divider()
        st.caption("Shared Dashboard")
