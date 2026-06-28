from __future__ import annotations

import pandas as pd

from repositories.dashboard_repository import load_action_trend_data, load_trend_data


def build_trend_dashboard_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    trend_df = load_trend_data()
    action_trend_df = load_action_trend_data()

    if trend_df.empty:
        return trend_df, action_trend_df

    trend_df = trend_df.copy()
    trend_df["Run"] = trend_df["run_id"].apply(lambda x: f"Run {x}")
    trend_df["Approval Rate"] = trend_df["approval_rate"].fillna(0)
    trend_df["Health Score"] = trend_df["health_score"].fillna(0)
    trend_df["Open Documents"] = trend_df["open_docs"].fillna(0)
    trend_df["Total Documents"] = trend_df["total_docs"].fillna(0)

    return trend_df, action_trend_df


def build_key_action_trends(action_trend_df: pd.DataFrame) -> pd.DataFrame:
    if action_trend_df is None or action_trend_df.empty:
        return pd.DataFrame()

    key_actions = [
        "OVERDUE / FOLLOW UP",
        "RETURNED BY NV5 / NEED RESUBMIT",
        "UPDATE TRACKING TO CLOSED",
    ]

    df = action_trend_df[action_trend_df["action"].isin(key_actions)].copy()

    if df.empty:
        return pd.DataFrame()

    pivot = df.pivot_table(
        index="run_id",
        columns="action",
        values="count",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    pivot["Run"] = pivot["run_id"].apply(lambda x: f"Run {x}")

    return pivot
