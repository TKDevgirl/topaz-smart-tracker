from __future__ import annotations

import streamlit as st

from config.settings import ADMIN_USERS


SESSION_DEFAULTS = {
    "logged_in": False,
    "role": "viewer",
    "username": "",
    "result_df": None,
    "report": None,
    "total_docs": 0,
    "open_docs": 0,
    "action_counts": {},
    "last_updated": "",
    "status_summary_df": None,
    "status_detail_df": None,
}


def init_session_state() -> None:
    for key, value in SESSION_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = value


def login(username: str) -> None:
    clean_username = username.strip()
    st.session_state.username = clean_username
    st.session_state.role = "admin" if clean_username.lower() in ADMIN_USERS else "viewer"
    st.session_state.logged_in = True


def logout() -> None:
    st.session_state.logged_in = False
    st.session_state.role = "viewer"
    st.session_state.username = ""
