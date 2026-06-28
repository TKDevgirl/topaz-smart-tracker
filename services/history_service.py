from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from config.settings import HISTORY_DIR


def save_history_snapshot(
    status_summary_df: pd.DataFrame | None,
    result_df: pd.DataFrame | None,
) -> str | None:
    if status_summary_df is None or status_summary_df.empty:
        return None

    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    summary_path = HISTORY_DIR / f"status_summary_{timestamp}.csv"
    status_summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")

    if result_df is not None and not result_df.empty:
        result_path = HISTORY_DIR / f"action_list_{timestamp}.csv"
        result_df.to_csv(result_path, index=False, encoding="utf-8-sig")

    return timestamp


def list_history_snapshots() -> pd.DataFrame:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)

    rows = []
    for path in sorted(HISTORY_DIR.glob("status_summary_*.csv"), reverse=True):
        timestamp = path.stem.replace("status_summary_", "")
        rows.append(
            {
                "Snapshot": timestamp,
                "File": str(path),
            }
        )

    return pd.DataFrame(rows)
