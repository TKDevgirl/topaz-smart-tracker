from __future__ import annotations

import streamlit as st

from config.settings import LOGO_PATH, PROJECT_NAME
from core.utils import image_to_base64


def render_logo_html(size_class: str = "hero-logo") -> str:
    logo64 = image_to_base64(LOGO_PATH)

    if logo64:
        return f'<img class="{size_class}" src="data:image/png;base64,{logo64}">'

    return '<div style="font-size:62px;">💎</div>'


def render_header() -> None:
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-grid">
                <div>{render_logo_html()}</div>
                <div>
                    <div class="hero-title">Topaz Smart Document Tracker</div>
                    <div class="hero-subtitle">{PROJECT_NAME} | Executive Document Control Dashboard</div>
                </div>
                <div class="hero-badge">V10 SQLite</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_role_message() -> None:
    if st.session_state.role == "admin":
        st.markdown(
            '<div class="info-box">👩‍💼 Admin mode: Admin can upload files, generate dashboard, and update shared data.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="viewer-box">ℹ️ Viewer mode: dashboard is read-only. Only Admin can upload files and update shared data.</div>',
            unsafe_allow_html=True,
        )
