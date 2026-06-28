import base64
import os
from io import BytesIO

import pandas as pd


def image_to_base64(path: str) -> str:
    if not os.path.exists(path):
        return ""

    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Summary") -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    output.seek(0)
    return output.getvalue()


def base_doc_no(doc_no):
    doc_no = str(doc_no).strip()
    parts = doc_no.split("-")

    if parts and parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])

    return doc_no
