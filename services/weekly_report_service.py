from __future__ import annotations

from io import BytesIO

import pandas as pd

from services.analytics_service import calculate_category_totals, calculate_project_health
from services.insight_service import generate_executive_insights, generate_recommended_actions
from services.trend_service import build_trend_dashboard_data


def build_weekly_report_excel(
    result_df: pd.DataFrame | None,
    status_summary_df: pd.DataFrame | None,
    action_counts: dict[str, int] | None,
) -> bytes:
    action_counts = action_counts or {}
    result_df = result_df if result_df is not None else pd.DataFrame()
    status_summary_df = status_summary_df if status_summary_df is not None else pd.DataFrame()

    health = calculate_project_health(status_summary_df, action_counts)
    category_df = calculate_category_totals(status_summary_df)
    trend_df, action_trend_df = build_trend_dashboard_data()

    insights_df = pd.DataFrame({"Executive Insights": generate_executive_insights(status_summary_df, action_counts)})
    actions_df = pd.DataFrame({"Recommended Actions": generate_recommended_actions(status_summary_df, action_counts)})

    kpi_df = pd.DataFrame(
        [
            {"Metric": "Project Health Score", "Value": health.get("score", 0)},
            {"Metric": "Project Health Level", "Value": health.get("level", "")},
            {"Metric": "Approval Rate", "Value": health.get("approved_rate", 0)},
            {"Metric": "Total Documents", "Value": health.get("total", 0)},
            {"Metric": "Approved", "Value": health.get("approved", 0)},
            {"Metric": "Open", "Value": health.get("open", 0)},
            {"Metric": "On Progress", "Value": health.get("on_progress", 0)},
            {"Metric": "Overdue", "Value": health.get("overdue", 0)},
            {"Metric": "Returned", "Value": health.get("returned", 0)},
            {"Metric": "Need Update", "Value": health.get("need_update", 0)},
        ]
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kpi_df.to_excel(writer, index=False, sheet_name="Executive KPI")
        insights_df.to_excel(writer, index=False, sheet_name="Insights")
        actions_df.to_excel(writer, index=False, sheet_name="Actions")
        status_summary_df.to_excel(writer, index=False, sheet_name="Status Summary")
        category_df.to_excel(writer, index=False, sheet_name="Category")
        trend_df.to_excel(writer, index=False, sheet_name="Trend")
        action_trend_df.to_excel(writer, index=False, sheet_name="Action Trend")
        result_df.to_excel(writer, index=False, sheet_name="Action List")

    output.seek(0)
    return output.getvalue()
