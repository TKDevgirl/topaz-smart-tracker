from __future__ import annotations

import streamlit as st

from repositories.dashboard_repository import list_dashboard_runs
from services.analytics_service import calculate_project_health
from services.enterprise_service import build_run_timeline, build_snapshot_compare, get_sync_summary


def render_sync_bar() -> None:
    status_summary_df = st.session_state.get("status_summary_df", None)
    action_counts = st.session_state.get("action_counts", {})
    health = calculate_project_health(status_summary_df, action_counts)

    summary = get_sync_summary(
        last_updated=st.session_state.get("last_updated", ""),
        total_docs=int(st.session_state.get("total_docs", 0)),
        health_score=health.get("score", 0),
    )

    st.markdown(
        f"""
        <div style="
            background:linear-gradient(90deg,#0f172a,#1e3a8a);
            color:white;
            border-radius:24px;
            padding:20px 24px;
            margin:16px 0 22px 0;
            box-shadow:0 14px 32px rgba(15,23,42,.18);">
            <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:18px;align-items:center;">
                <div>
                    <div style="font-size:12px;color:#bfdbfe;font-weight:800;">Platform</div>
                    <div style="font-size:22px;font-weight:950;">V11 Enterprise</div>
                </div>
                <div>
                    <div style="font-size:12px;color:#bfdbfe;font-weight:800;">SQLite</div>
                    <div style="font-size:20px;font-weight:900;">{summary['sqlite']}</div>
                </div>
                <div>
                    <div style="font-size:12px;color:#bfdbfe;font-weight:800;">Last Sync</div>
                    <div style="font-size:18px;font-weight:900;">{summary['last_sync']}</div>
                </div>
                <div>
                    <div style="font-size:12px;color:#bfdbfe;font-weight:800;">Documents</div>
                    <div style="font-size:24px;font-weight:950;">{summary['total_docs']}</div>
                </div>
                <div>
                    <div style="font-size:12px;color:#bfdbfe;font-weight:800;">Health</div>
                    <div style="font-size:20px;font-weight:950;">{summary['health']}</div>
                    <div style="font-size:12px;color:#dbeafe;">{summary['score']}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_alert_center() -> None:
    action_counts = st.session_state.get("action_counts", {})

    overdue = int(action_counts.get("OVERDUE / FOLLOW UP", 0))
    returned = int(action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0))
    need_update = int(action_counts.get("UPDATE TRACKING TO CLOSED", 0))

    st.markdown(
        '<div class="panel"><div class="panel-title">🚨 Alert Center</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        if overdue > 0:
            st.error(f"🔴 Critical: {overdue} overdue item(s)")
        else:
            st.success("🟢 No overdue items")

    with c2:
        if returned > 5:
            st.warning(f"🟡 Warning: {returned} returned item(s)")
        elif returned > 0:
            st.info(f"🔵 Returned: {returned} item(s)")
        else:
            st.success("🟢 No returned items")

    with c3:
        if need_update > 0:
            st.info(f"🔵 Need Update: {need_update} item(s)")
        else:
            st.success("🟢 No update required")

    st.markdown("</div>", unsafe_allow_html=True)


def render_run_timeline_panel() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">🕘 Executive Run Timeline</div>',
        unsafe_allow_html=True,
    )

    runs_df = build_run_timeline()

    if runs_df.empty:
        st.info("No SQLite runs available yet.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    display_cols = [
        col for col in [
            "Run ID",
            "Run Time",
            "Uploaded By",
            "Total Documents",
            "Open Documents",
            "Approval Rate",
            "Health Score",
        ]
        if col in runs_df.columns
    ]

    st.dataframe(runs_df[display_cols], use_container_width=True, hide_index=True, height=230)

    latest_run = runs_df.iloc[0]
    st.caption(
        f"Latest run: #{latest_run.get('Run ID', '-')} | {latest_run.get('Run Time', '-')}"
    )

    st.markdown("</div>", unsafe_allow_html=True)


def render_snapshot_compare_panel() -> None:
    st.markdown(
        '<div class="panel"><div class="panel-title">🔁 Snapshot Compare</div>',
        unsafe_allow_html=True,
    )

    runs_df = list_dashboard_runs()

    if runs_df.empty or len(runs_df) < 2:
        st.info("Snapshot Compare will be available after at least 2 dashboard generations.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    run_ids = runs_df["Run ID"].tolist()

    c1, c2 = st.columns(2)

    with c1:
        from_run = st.selectbox("From Run", run_ids, index=min(1, len(run_ids)-1), key="compare_from_run")

    with c2:
        to_run = st.selectbox("To Run", run_ids, index=0, key="compare_to_run")

    compare_df = build_snapshot_compare(int(from_run), int(to_run))

    if compare_df.empty:
        st.info("No compare data found.")
    else:
        st.dataframe(compare_df, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)
