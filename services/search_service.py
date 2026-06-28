from __future__ import annotations

import pandas as pd


def search_documents(
    detail_df: pd.DataFrame | None,
    query: str = "",
    document_type: str = "All",
    category: str = "All",
    status: str = "All",
) -> pd.DataFrame:
    """
    Smart search over latest document detail data.
    Searches across Document No, Document Name, Category, Status, Info, Revision.
    """
    if detail_df is None or detail_df.empty:
        return pd.DataFrame()

    df = detail_df.copy()

    if document_type != "All" and "Tracking Sheet" in df.columns:
        df = df[df["Tracking Sheet"] == document_type]

    if category != "All" and "Category" in df.columns:
        df = df[df["Category"] == category]

    if status != "All" and "Status" in df.columns:
        df = df[df["Status"] == status]

    query = str(query or "").strip()

    if query:
        search_cols = [
            col for col in [
                "Document No",
                "Document Name",
                "Category",
                "Status",
                "Info",
                "Revision",
                "Tracking Sheet",
            ]
            if col in df.columns
        ]

        mask = df[search_cols].astype(str).apply(
            lambda row: row.str.contains(query, case=False, na=False).any(),
            axis=1,
        )

        df = df[mask]

    preferred_cols = [
        "Tracking Sheet",
        "Document No",
        "Revision",
        "Category",
        "Document Name",
        "Status",
        "Info",
    ]

    cols = [col for col in preferred_cols if col in df.columns]
    other_cols = [col for col in df.columns if col not in cols]

    return df[cols + other_cols]


def get_filter_options(detail_df: pd.DataFrame | None) -> dict[str, list[str]]:
    if detail_df is None or detail_df.empty:
        return {
            "document_types": ["All"],
            "categories": ["All"],
            "statuses": ["All"],
        }

    document_types = ["All"]
    categories = ["All"]
    statuses = ["All"]

    if "Tracking Sheet" in detail_df.columns:
        document_types += sorted(detail_df["Tracking Sheet"].dropna().astype(str).unique().tolist())

    if "Category" in detail_df.columns:
        categories += sorted(detail_df["Category"].dropna().astype(str).unique().tolist())

    if "Status" in detail_df.columns:
        statuses += sorted(detail_df["Status"].dropna().astype(str).unique().tolist())

    return {
        "document_types": document_types,
        "categories": categories,
        "statuses": statuses,
    }


def build_document_timeline(detail_df: pd.DataFrame | None, document_no: str) -> pd.DataFrame:
    """
    V8.1 timeline uses available latest/detail data.
    If full revision history is loaded in future, this function can support all revisions.
    """
    if detail_df is None or detail_df.empty or not document_no:
        return pd.DataFrame()

    df = detail_df.copy()

    if "Document No" not in df.columns:
        return pd.DataFrame()

    timeline_df = df[df["Document No"].astype(str) == str(document_no)].copy()

    preferred_cols = [
        "Tracking Sheet",
        "Document No",
        "Revision",
        "Category",
        "Status",
        "Document Name",
        "Info",
    ]

    cols = [col for col in preferred_cols if col in timeline_df.columns]
    other_cols = [col for col in timeline_df.columns if col not in cols]

    return timeline_df[cols + other_cols]
