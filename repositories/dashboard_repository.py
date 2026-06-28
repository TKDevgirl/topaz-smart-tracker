from __future__ import annotations

from datetime import datetime

import pandas as pd

from core.database import get_connection, get_latest_run_id, init_database


def _clean_value(value):
    if pd.isna(value):
        return ""
    return str(value)


def save_dashboard_run(
    result_df: pd.DataFrame,
    status_summary_df: pd.DataFrame,
    status_detail_df: pd.DataFrame,
    total_docs: int,
    open_docs: int,
    approval_rate: float,
    health_score: float,
    uploaded_by: str,
) -> int:
    init_database()
    conn = get_connection()
    cur = conn.cursor()

    run_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

    cur.execute(
        """
        INSERT INTO dashboard_runs (
            run_time, uploaded_by, total_docs, open_docs, approval_rate, health_score
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (run_time, uploaded_by, int(total_docs), int(open_docs), float(approval_rate), float(health_score)),
    )

    run_id = int(cur.lastrowid)

    if result_df is not None and not result_df.empty:
        for _, row in result_df.iterrows():
            cur.execute(
                """
                INSERT INTO action_list (
                    run_id,
                    tracking_sheet,
                    document_no,
                    document_name,
                    tracking_status,
                    info,
                    takenaka_status_1,
                    takenaka_status_2,
                    takenaka_status_3,
                    action,
                    checked_time
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    _clean_value(row.get("Tracking Sheet", "")),
                    _clean_value(row.get("Document No", "")),
                    _clean_value(row.get("Document Name", "")),
                    _clean_value(row.get("Tracking Status", "")),
                    _clean_value(row.get("Info", "")),
                    _clean_value(row.get("Takenaka Status 1", "")),
                    _clean_value(row.get("Takenaka Status 2", "")),
                    _clean_value(row.get("Takenaka Status 3", "")),
                    _clean_value(row.get("Action", "")),
                    _clean_value(row.get("Checked Time", "")),
                ),
            )

    if status_summary_df is not None and not status_summary_df.empty:
        for _, row in status_summary_df.iterrows():
            cur.execute(
                """
                INSERT INTO status_summary (
                    run_id,
                    document_type,
                    status,
                    total,
                    mat,
                    mcr,
                    mts,
                    cvi,
                    dwg
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    _clean_value(row.get("Document Type", "")),
                    _clean_value(row.get("Status", "")),
                    _clean_value(row.get("Total", "")),
                    _clean_value(row.get("MAT", "")),
                    _clean_value(row.get("MCR", "")),
                    _clean_value(row.get("MTS", "")),
                    _clean_value(row.get("CVI", "")),
                    _clean_value(row.get("DWG", "")),
                ),
            )

    if status_detail_df is not None and not status_detail_df.empty:
        for _, row in status_detail_df.iterrows():
            cur.execute(
                """
                INSERT INTO status_detail (
                    run_id,
                    tracking_sheet,
                    document_no,
                    category,
                    document_name,
                    revision,
                    status,
                    info
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    _clean_value(row.get("Tracking Sheet", "")),
                    _clean_value(row.get("Document No", "")),
                    _clean_value(row.get("Category", "")),
                    _clean_value(row.get("Document Name", "")),
                    _clean_value(row.get("Revision", "")),
                    _clean_value(row.get("Status", "")),
                    _clean_value(row.get("Info", "")),
                ),
            )

    conn.commit()
    conn.close()
    return run_id


def load_dashboard_run(run_id: int | None = None):
    init_database()

    if run_id is None:
        run_id = get_latest_run_id()

    if run_id is None:
        return None, None, None, None

    conn = get_connection()

    meta_df = pd.read_sql_query(
        "SELECT * FROM dashboard_runs WHERE id = ?",
        conn,
        params=(run_id,),
    )

    if meta_df.empty:
        conn.close()
        return None, None, None, None

    result_df = pd.read_sql_query(
        """
        SELECT
            tracking_sheet AS 'Tracking Sheet',
            document_no AS 'Document No',
            document_name AS 'Document Name',
            tracking_status AS 'Tracking Status',
            info AS 'Info',
            takenaka_status_1 AS 'Takenaka Status 1',
            takenaka_status_2 AS 'Takenaka Status 2',
            takenaka_status_3 AS 'Takenaka Status 3',
            action AS 'Action',
            checked_time AS 'Checked Time'
        FROM action_list
        WHERE run_id = ?
        ORDER BY id
        """,
        conn,
        params=(run_id,),
    )

    status_summary_df = pd.read_sql_query(
        """
        SELECT
            document_type AS 'Document Type',
            status AS 'Status',
            total AS 'Total',
            mat AS 'MAT',
            mcr AS 'MCR',
            mts AS 'MTS',
            cvi AS 'CVI',
            dwg AS 'DWG'
        FROM status_summary
        WHERE run_id = ?
        ORDER BY id
        """,
        conn,
        params=(run_id,),
    )

    status_detail_df = pd.read_sql_query(
        """
        SELECT
            tracking_sheet AS 'Tracking Sheet',
            document_no AS 'Document No',
            category AS 'Category',
            document_name AS 'Document Name',
            revision AS 'Revision',
            status AS 'Status',
            info AS 'Info'
        FROM status_detail
        WHERE run_id = ?
        ORDER BY id
        """,
        conn,
        params=(run_id,),
    )

    conn.close()

    meta = meta_df.iloc[0].to_dict()
    return meta, result_df, status_summary_df, status_detail_df


def list_dashboard_runs() -> pd.DataFrame:
    init_database()
    conn = get_connection()

    runs_df = pd.read_sql_query(
        """
        SELECT
            id AS 'Run ID',
            run_time AS 'Run Time',
            uploaded_by AS 'Uploaded By',
            total_docs AS 'Total Documents',
            open_docs AS 'Open Documents',
            approval_rate AS 'Approval Rate',
            health_score AS 'Health Score'
        FROM dashboard_runs
        ORDER BY id DESC
        """,
        conn,
    )

    conn.close()
    return runs_df


def load_trend_data() -> pd.DataFrame:
    """
    Load run-level trend data from SQLite.
    """
    init_database()
    conn = get_connection()

    trend_df = pd.read_sql_query(
        """
        SELECT
            id AS run_id,
            run_time,
            total_docs,
            open_docs,
            approval_rate,
            health_score
        FROM dashboard_runs
        ORDER BY id
        """,
        conn,
    )

    conn.close()
    return trend_df


def load_action_trend_data() -> pd.DataFrame:
    """
    Load action counts by run_id for trend dashboard.
    """
    init_database()
    conn = get_connection()

    df = pd.read_sql_query(
        """
        SELECT
            run_id,
            action,
            COUNT(*) AS count
        FROM action_list
        GROUP BY run_id, action
        ORDER BY run_id
        """,
        conn,
    )

    conn.close()
    return df
