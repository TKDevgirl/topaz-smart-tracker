from __future__ import annotations

import pandas as pd

from config.columns import STANDARD_CATEGORIES, STANDARD_STATUSES
from services.tracking_service import get_latest_revision_records, read_tracking_records


def build_one_summary(df: pd.DataFrame, document_type: str, categories: list[str]) -> pd.DataFrame:
    df = df[df["Status"].isin(STANDARD_STATUSES)].copy()

    rows: list[dict[str, object]] = []

    for status in STANDARD_STATUSES:
        status_df = df[df["Status"] == status]
        row: dict[str, object] = {
            "Document Type": document_type,
            "Status": status,
            "Total": len(status_df),
        }

        for category in categories:
            row[category] = len(status_df[status_df["Category"] == category])

        rows.append(row)

    total_row: dict[str, object] = {
        "Document Type": document_type,
        "Status": "Total Document",
        "Total": len(df),
    }

    for category in categories:
        total_row[category] = len(df[df["Category"] == category])

    rows.append(total_row)

    summary_df = pd.DataFrame(rows)
    display_df = summary_df.copy()

    for col in display_df.columns:
        if col not in ["Document Type", "Status"]:
            display_df[col] = display_df[col].apply(lambda x: "-" if int(x) == 0 else int(x))

    return display_df


def build_status_summary_from_tracking(tracking_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    records_df = read_tracking_records(tracking_file)

    if records_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    latest_df = get_latest_revision_records(records_df)
    categories = STANDARD_CATEGORIES.copy()

    rfa_df = latest_df[latest_df["Tracking Sheet"] == "RFA"].copy()
    rfi_df = latest_df[latest_df["Tracking Sheet"] == "RFI"].copy()

    rfa_summary = build_one_summary(rfa_df, "RFA", categories)
    rfi_summary = build_one_summary(rfi_df, "RFI", categories)

    summary_cols = ["Document Type", "Status", "Total", "MAT", "MCR", "MTS", "CVI", "DWG"]
    rfa_summary = rfa_summary[summary_cols]
    rfi_summary = rfi_summary[summary_cols]

    combined_summary = pd.concat([rfa_summary, rfi_summary], ignore_index=True)

    latest_detail_df = latest_df.drop(columns=["Revision Sort", "Excel Row"], errors="ignore")

    return combined_summary, latest_detail_df
