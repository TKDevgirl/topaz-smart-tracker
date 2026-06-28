from __future__ import annotations

import pandas as pd

from services.analytics_service import calculate_category_totals, calculate_project_health, calculate_status_totals


def generate_executive_insights(
    status_summary_df: pd.DataFrame | None,
    action_counts: dict[str, int] | None,
) -> list[str]:
    action_counts = action_counts or {}
    health = calculate_project_health(status_summary_df, action_counts)
    totals = calculate_status_totals(status_summary_df)
    category_df = calculate_category_totals(status_summary_df)

    insights: list[str] = []

    insights.append(
        f"{health['emoji']} Project health is {health['level']} with a score of {health['score']}/100."
    )

    if totals["total"] > 0:
        insights.append(
            f"Overall approval rate is {health['approved_rate']}% "
            f"({totals['approved']} approved out of {totals['total']} latest documents)."
        )

    if totals["open"] > 0:
        insights.append(
            f"There are {totals['open']} open documents that may require follow-up."
        )

    if totals["on_progress"] > 0:
        insights.append(
            f"{totals['on_progress']} documents are currently on progress."
        )

    overdue = int(action_counts.get("OVERDUE / FOLLOW UP", 0))
    returned = int(action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0))
    need_update = int(action_counts.get("UPDATE TRACKING TO CLOSED", 0))

    if overdue > 0:
        insights.append(
            f"Overdue risk detected: {overdue} item(s) require urgent follow-up."
        )

    if returned > 0:
        insights.append(
            f"{returned} item(s) were returned by NV5 and may need resubmission."
        )

    if need_update > 0:
        insights.append(
            f"{need_update} item(s) appear ready to update tracking status to closed."
        )

    if category_df is not None and not category_df.empty:
        category_df = category_df.sort_values("Count", ascending=False)
        top = category_df.iloc[0]

        if int(top["Count"]) > 0:
            insights.append(
                f"Highest document volume is in {top['Category']} with {int(top['Count'])} document(s)."
            )

    if len(insights) <= 1:
        insights.append("No major risk detected from the current dashboard data.")

    return insights


def generate_recommended_actions(
    status_summary_df: pd.DataFrame | None,
    action_counts: dict[str, int] | None,
) -> list[str]:
    action_counts = action_counts or {}
    health = calculate_project_health(status_summary_df, action_counts)

    actions: list[str] = []

    if health["overdue"] > 0:
        actions.append("Prioritize overdue documents and assign owners for follow-up.")

    if health["returned"] > 0:
        actions.append("Review returned documents and prepare resubmission plan.")

    if health["need_update"] > 0:
        actions.append("Update tracking records for documents already closed in Takenaka.")

    if health["open"] > 0:
        actions.append("Review open documents and confirm next action with responsible parties.")

    if health["score"] >= 85:
        actions.append("Maintain current tracking process and continue weekly monitoring.")

    if not actions:
        actions.append("No immediate action required. Continue monitoring the dashboard.")

    return actions
