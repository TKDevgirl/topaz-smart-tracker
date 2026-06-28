from __future__ import annotations

import pandas as pd

from repositories.dashboard_repository import compare_runs, list_dashboard_runs


def get_sync_summary(last_updated: str, total_docs: int, health_score: float | int | str) -> dict[str, str]:
    try:
        score = float(health_score)
    except Exception:
        score = 0.0

    if score >= 85:
        status = "Healthy"
        emoji = "🟢"
    elif score >= 65:
        status = "Watch"
        emoji = "🟡"
    else:
        status = "Risk"
        emoji = "🔴"

    return {
        "sqlite": "Connected",
        "last_sync": last_updated or "-",
        "total_docs": str(total_docs),
        "health": f"{emoji} {status}",
        "score": f"{score:.1f}/100",
    }


def build_run_timeline() -> pd.DataFrame:
    runs_df = list_dashboard_runs()

    if runs_df.empty:
        return runs_df

    return runs_df.copy()


def build_snapshot_compare(from_run_id: int, to_run_id: int) -> pd.DataFrame:
    return compare_runs(from_run_id, to_run_id)
