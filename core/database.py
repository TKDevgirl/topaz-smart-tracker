from __future__ import annotations

import sqlite3
from pathlib import Path

from config.settings import DATABASE_PATH


def get_connection() -> sqlite3.Connection:
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_database() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS dashboard_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_time TEXT NOT NULL,
            uploaded_by TEXT,
            total_docs INTEGER,
            open_docs INTEGER,
            approval_rate REAL,
            health_score REAL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS action_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            tracking_sheet TEXT,
            document_no TEXT,
            document_name TEXT,
            tracking_status TEXT,
            info TEXT,
            takenaka_status_1 TEXT,
            takenaka_status_2 TEXT,
            takenaka_status_3 TEXT,
            action TEXT,
            checked_time TEXT,
            FOREIGN KEY(run_id) REFERENCES dashboard_runs(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS status_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            document_type TEXT,
            status TEXT,
            total TEXT,
            mat TEXT,
            mcr TEXT,
            mts TEXT,
            cvi TEXT,
            dwg TEXT,
            FOREIGN KEY(run_id) REFERENCES dashboard_runs(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS status_detail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            tracking_sheet TEXT,
            document_no TEXT,
            category TEXT,
            document_name TEXT,
            revision TEXT,
            status TEXT,
            info TEXT,
            FOREIGN KEY(run_id) REFERENCES dashboard_runs(id)
        )
        """
    )

    conn.commit()
    conn.close()


def get_latest_run_id() -> int | None:
    init_database()
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id FROM dashboard_runs ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    conn.close()

    if row is None:
        return None

    return int(row["id"])
