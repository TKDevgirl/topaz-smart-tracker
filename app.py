import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================================
# Topaz Smart Tracker V6.1.1
# Focus: Raw Status Summary from RFA / RFI only
# =========================================================

st.set_page_config(
    page_title="Topaz Smart Tracker V6.1 Enterprise",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.sidebar.title("📊 Topaz Tracker")
st.sidebar.caption("V6.1 Enterprise")

uploaded_file = st.sidebar.file_uploader(
    "Upload Tracker File",
    type=["xlsx", "xls"],
    help="Upload Tracking_document Excel file."
)

st.title("📊 Topaz Smart Tracker V6.1 Enterprise")
st.caption("Document Control Dashboard for RFI / RFA / Material / Drawing / Schedule Tracking")


def normalize_status(value):
    value = str(value).strip()

    mapping = {
        "open": "Open",
        "on progress": "On Progress",
        "in progress": "On Progress",
        "approved": "Approved",
        "approve": "Approved",
        "close": "Close",
        "closed": "Close",
    }

    return mapping.get(value.lower(), value)


def load_rfa_rfi(file):
    frames = []

    for sheet_name in ["RFA", "RFI"]:
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name)
        except Exception:
            continue

        if df.empty:
            continue

        # Required columns from the real Tracking Document
        if "Category" not in df.columns or "Status" not in df.columns:
            continue

        temp = df[["Category", "Status"]].copy()
        temp["Sheet"] = sheet_name
        temp["Category"] = temp["Category"].astype(str).str.strip()
        temp["Status"] = temp["Status"].apply(normalize_status)

        temp = temp[
            (temp["Category"] != "") &
            (temp["Category"].str.lower() != "nan") &
            (temp["Status"] != "") &
            (temp["Status"].str.lower() != "nan")
        ]

        frames.append(temp)

    if not frames:
        return pd.DataFrame(columns=["Category", "Status", "Sheet"])

    return pd.concat(frames, ignore_index=True)


def build_status_summary(df):
    if df.empty:
        return pd.DataFrame()

    # Show only what user asked first
    keep_status = ["Open", "On Progress", "Approved"]
    df_keep = df[df["Status"].isin(keep_status)].copy()

    categories = sorted(df_keep["Category"].dropna().unique().tolist())

    rows = []

    total_row = {"Status": "Total Document"}
    total_row["Total"] = len(df_keep)
    for cat in categories:
        total_row[cat] = len(df_keep[df_keep["Category"] == cat])
    rows.append(total_row)

    for status in keep_status:
        row = {"Status": status}
        status_df = df_keep[df_keep["Status"] == status]
        row["Total"] = len(status_df)

        for cat in categories:
            row[cat] = len(status_df[status_df["Category"] == cat])

        rows.append(row)

    summary = pd.DataFrame(rows)

    # Preferred display order
    preferred_cols = ["Status", "Total", "MAT", "MCR", "MTS", "CVI", "DWG"]
    final_cols = [c for c in preferred_cols if c in summary.columns]
    other_cols = [c for c in summary.columns if c not in final_cols]

    return summary[final_cols + other_cols]


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")
    return output.getvalue()


if uploaded_file is None:
    st.info("Please upload Tracking_document Excel file to show Raw Data Status Summary from Sheet RFA / RFI.")
    st.stop()


raw_df = load_rfa_rfi(uploaded_file)

if raw_df.empty:
    st.error("Cannot find valid data. Please check that your Excel file has Sheet RFA / RFI and columns Category + Status.")
    st.stop()


summary_df = build_status_summary(raw_df)

# KPI
total_document = int(summary_df.loc[summary_df["Status"] == "Total Document", "Total"].iloc[0])
open_count = int(summary_df.loc[summary_df["Status"] == "Open", "Total"].iloc[0])
progress_count = int(summary_df.loc[summary_df["Status"] == "On Progress", "Total"].iloc[0])
approved_count = int(summary_df.loc[summary_df["Status"] == "Approved", "Total"].iloc[0])

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Document", total_document)
k2.metric("Open", open_count)
k3.metric("On Progress", progress_count)
k4.metric("Approved", approved_count)

st.divider()

st.subheader("📌 Raw Data Status Summary")
st.dataframe(summary_df, use_container_width=True, hide_index=True)

st.download_button(
    label="Download Summary",
    data=to_excel_bytes(summary_df),
    file_name="Topaz_Raw_Status_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("View Raw Data from RFA / RFI"):
    st.dataframe(raw_df, use_container_width=True, hide_index=True)

st.caption("V6.1.1 | Summary reads directly from Sheet RFA and RFI using Category and Status.")
