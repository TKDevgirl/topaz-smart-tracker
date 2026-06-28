from __future__ import annotations

from io import BytesIO

import pandas as pd


def build_database_backup_excel(
    runs_df: pd.DataFrame | None,
    result_df: pd.DataFrame | None,
    status_summary_df: pd.DataFrame | None,
    status_detail_df: pd.DataFrame | None,
) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        (runs_df if runs_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="Runs")
        (status_summary_df if status_summary_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="Status Summary")
        (status_detail_df if status_detail_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="Status Detail")
        (result_df if result_df is not None else pd.DataFrame()).to_excel(writer, index=False, sheet_name="Action List")

    output.seek(0)
    return output.getvalue()
