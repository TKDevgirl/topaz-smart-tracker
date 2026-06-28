from __future__ import annotations

import pandas as pd


def safe_int(value) -> int:
    try:
        if value == "-":
            return 0
        return int(value)
    except Exception:
        return 0


def calculate_status_totals(status_summary_df: pd.DataFrame | None) -> dict[str, int]:
    if status_summary_df is None or status_summary_df.empty:
        return {
            "total": 0,
            "open": 0,
            "on_progress": 0,
            "approved": 0,
        }

    totals = {
        "total": 0,
        "open": 0,
        "on_progress": 0,
        "approved": 0,
    }

    for _, row in status_summary_df.iterrows():
        status = str(row.get("Status", "")).strip()
        value = safe_int(row.get("Total", 0))

        if status == "Total Document":
            totals["total"] += value
        elif status == "Open":
            totals["open"] += value
        elif status == "On Progress":
            totals["on_progress"] += value
        elif status == "Approved":
            totals["approved"] += value

    return totals


def calculate_category_totals(status_summary_df: pd.DataFrame | None) -> pd.DataFrame:
    if status_summary_df is None or status_summary_df.empty:
        return pd.DataFrame(columns=["Category", "Count"])

    categories = ["MAT", "MCR", "MTS", "CVI", "DWG"]
    rows = []

    total_rows = status_summary_df[status_summary_df["Status"] == "Total Document"]

    for category in categories:
        count = 0

        if category in total_rows.columns:
            for value in total_rows[category].tolist():
                count += safe_int(value)

        rows.append({"Category": category, "Count": count})

    return pd.DataFrame(rows)


def calculate_project_health(
    status_summary_df: pd.DataFrame | None,
    action_counts: dict[str, int] | None,
) -> dict[str, object]:
    action_counts = action_counts or {}
    totals = calculate_status_totals(status_summary_df)

    total = totals["total"]
    approved = totals["approved"]
    open_count = totals["open"]
    progress = totals["on_progress"]

    overdue = int(action_counts.get("OVERDUE / FOLLOW UP", 0))
    returned = int(action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0))
    need_update = int(action_counts.get("UPDATE TRACKING TO CLOSED", 0))

    if total == 0:
        score = 0
        approved_rate = 0
    else:
        approved_rate = round((approved / total) * 100, 1)
        open_penalty = (open_count / total) * 20
        progress_penalty = (progress / total) * 8
        overdue_penalty = overdue * 5
        returned_penalty = returned * 3
        update_penalty = need_update * 1

        score = round(
            max(
                0,
                min(
                    100,
                    100 - open_penalty - progress_penalty - overdue_penalty - returned_penalty - update_penalty,
                ),
            ),
            1,
        )

    if score >= 85:
        level = "Healthy"
        emoji = "🟢"
    elif score >= 65:
        level = "Watch"
        emoji = "🟡"
    else:
        level = "Risk"
        emoji = "🔴"

    return {
        "score": score,
        "level": level,
        "emoji": emoji,
        "approved_rate": approved_rate,
        "total": total,
        "approved": approved,
        "open": open_count,
        "on_progress": progress,
        "overdue": overdue,
        "returned": returned,
        "need_update": need_update,
    }
