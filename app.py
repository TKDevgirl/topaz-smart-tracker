import streamlit as st
import pandas as pd
from datetime import date, datetime
from io import BytesIO

# =========================================================
# Topaz Smart Tracker - V6.1 Enterprise
# Author: Pami / ByteBridge
# Purpose: Document / RFI / RFA / Material / Drawing tracker
# =========================================================

st.set_page_config(
    page_title="Topaz Smart Tracker V6.1 Enterprise",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------
# Standard Columns
# -----------------------------
STANDARD_COLUMNS = [
    "No",
    "Document ID",
    "Category",
    "Title / Description",
    "Discipline",
    "Rev",
    "Status",
    "Action By",
    "Submission Date",
    "Due Date",
    "Closed Date",
    "Aconex / Link",
    "Remark"
]

STATUS_OPTIONS = [
    "Draft",
    "Open",
    "Submitted",
    "Under Review",
    "Approved",
    "Approved with Comment",
    "Rejected",
    "Closed",
    "Overdue"
]

CATEGORY_OPTIONS = [
    "Tender",
    "Drawing",
    "RFI",
    "RFA",
    "Material",
    "Schedule",
    "BOM",
    "Installation",
    "Commissioning",
    "Weekly",
    "Other"
]

# -----------------------------
# Helper Functions
# -----------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure required columns exist."""
    df = df.copy()

    for col in STANDARD_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[STANDARD_COLUMNS]

    # Convert dates safely
    for col in ["Submission Date", "Due Date", "Closed Date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

    # Clean status/category
    df["Status"] = df["Status"].fillna("Open").astype(str).str.strip()
    df["Category"] = df["Category"].fillna("Other").astype(str).str.strip()

    # Auto No
    if df["No"].isna().all() or (df["No"].astype(str).str.strip() == "").all():
        df["No"] = range(1, len(df) + 1)

    return df


def create_template() -> pd.DataFrame:
    data = [
        [1, "RFI-001", "RFI", "Clarification for outdoor fiber backbone scope", "ICT", "00", "Open", "NV5 / DE", date.today(), "", "", "", "Need confirm demarcation and responsible party"],
        [2, "RFA-MAT-001", "Material", "Fiber tray material submission", "ICT", "00", "Under Review", "NV5", date.today(), "", "", "", "NXTGEN FRP tray"],
        [3, "DWG-001", "Drawing", "ODF rack layout drawing", "ICT", "00", "Draft", "ByteBridge", date.today(), "", "", "", "Prepare for submission"],
    ]
    return pd.DataFrame(data, columns=STANDARD_COLUMNS)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return create_template()

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file)
    else:
        st.error("Please upload only .xlsx, .xls or .csv file.")
        return create_template()

    return normalize_columns(df)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tracker")
    return output.getvalue()


def status_badge(status: str) -> str:
    status = str(status).strip()
    icon_map = {
        "Draft": "⚪ Draft",
        "Open": "🔵 Open",
        "Submitted": "🟣 Submitted",
        "Under Review": "🟡 Under Review",
        "Approved": "🟢 Approved",
        "Approved with Comment": "🟢 Approved with Comment",
        "Rejected": "🔴 Rejected",
        "Closed": "✅ Closed",
        "Overdue": "⛔ Overdue"
    }
    return icon_map.get(status, status)


def add_overdue_flag(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    today = date.today()

    def check(row):
        due = row.get("Due Date")
        status = str(row.get("Status", "")).strip()

        if pd.isna(due) or due == "":
            return status

        if status in ["Approved", "Approved with Comment", "Closed", "Rejected"]:
            return status

        if due < today:
            return "Overdue"

        return status

    df["Calculated Status"] = df.apply(check, axis=1)
    df["Status Display"] = df["Calculated Status"].apply(status_badge)
    return df


# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("📊 Topaz Tracker")
st.sidebar.caption("V6.1 Enterprise")

uploaded_file = st.sidebar.file_uploader(
    "Upload Tracker File",
    type=["xlsx", "xls", "csv"],
    help="Upload your Master Document Register / RFI / RFA tracker."
)

df_raw = read_uploaded_file(uploaded_file)
df = normalize_columns(df_raw)
df = add_overdue_flag(df)

st.sidebar.divider()

st.sidebar.subheader("Filters")

category_filter = st.sidebar.multiselect(
    "Category",
    options=sorted(df["Category"].dropna().unique().tolist()),
    default=sorted(df["Category"].dropna().unique().tolist())
)

status_filter = st.sidebar.multiselect(
    "Status",
    options=sorted(df["Calculated Status"].dropna().unique().tolist()),
    default=sorted(df["Calculated Status"].dropna().unique().tolist())
)

discipline_filter = st.sidebar.multiselect(
    "Discipline",
    options=sorted(df["Discipline"].dropna().astype(str).unique().tolist()),
    default=sorted(df["Discipline"].dropna().astype(str).unique().tolist())
)

search_text = st.sidebar.text_input("Search", placeholder="Document ID / Title / Remark")

# Apply filters
filtered_df = df.copy()

if category_filter:
    filtered_df = filtered_df[filtered_df["Category"].isin(category_filter)]

if status_filter:
    filtered_df = filtered_df[filtered_df["Calculated Status"].isin(status_filter)]

if discipline_filter:
    filtered_df = filtered_df[filtered_df["Discipline"].astype(str).isin(discipline_filter)]

if search_text:
    search_text_lower = search_text.lower()
    filtered_df = filtered_df[
        filtered_df.astype(str).apply(
            lambda row: row.str.lower().str.contains(search_text_lower, na=False).any(),
            axis=1
        )
    ]

# -----------------------------
# Main Header
# -----------------------------
st.title("📊 Topaz Smart Tracker V6.1 Enterprise")
st.caption("Document Control Dashboard for RFI / RFA / Material / Drawing / Schedule Tracking")

if uploaded_file is None:
    st.info("No file uploaded yet. Showing sample template data. You can download the template below and use it as your tracker.")

# -----------------------------
# KPI Cards
# -----------------------------
total_docs = len(filtered_df)
open_docs = len(filtered_df[filtered_df["Calculated Status"].isin(["Open", "Submitted", "Under Review", "Draft"])])
approved_docs = len(filtered_df[filtered_df["Calculated Status"].isin(["Approved", "Approved with Comment"])])
overdue_docs = len(filtered_df[filtered_df["Calculated Status"] == "Overdue"])
closed_docs = len(filtered_df[filtered_df["Calculated Status"] == "Closed"])

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total", total_docs)
c2.metric("Open / Active", open_docs)
c3.metric("Approved", approved_docs)
c4.metric("Overdue", overdue_docs)
c5.metric("Closed", closed_docs)

st.divider()

# -----------------------------
# Charts
# -----------------------------
left, right = st.columns(2)

with left:
    st.subheader("Status Summary")
    status_chart = (
        filtered_df["Calculated Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
    )
    st.bar_chart(status_chart, x="Status", y="Count", use_container_width=True)

with right:
    st.subheader("Category Summary")
    category_chart = (
        filtered_df["Category"]
        .value_counts()
        .rename_axis("Category")
        .reset_index(name="Count")
    )
    st.bar_chart(category_chart, x="Category", y="Count", use_container_width=True)



# -----------------------------
# Raw Data Summary from RFA / RFI
# -----------------------------
st.subheader("📌 Raw Data Status Summary")

def build_raw_status_summary(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()

    summary_frames = []

    for sheet in ["RFA", "RFI"]:
        try:
            raw = pd.read_excel(uploaded_file, sheet_name=sheet)
        except Exception:
            continue

        if raw.empty or raw.shape[1] < 6:
            continue

        # Column F = Status
        status_col = raw.columns[5]

        # Try to detect category column
        possible_category_cols = [
            col for col in raw.columns
            if str(col).strip().lower() in [
                "category", "type", "document type", "doc type", "submission type"
            ]
        ]

        if possible_category_cols:
            category_col = possible_category_cols[0]
            temp = raw[[category_col, status_col]].copy()
            temp.columns = ["Category", "Status"]
        else:
            # If category column cannot be found, use sheet name as category
            temp = raw[[status_col]].copy()
            temp.columns = ["Status"]
            temp["Category"] = sheet

        temp["Status"] = temp["Status"].astype(str).str.strip()
        temp["Category"] = temp["Category"].astype(str).str.strip()

        temp = temp[
            (temp["Status"] != "") &
            (temp["Status"].str.lower() != "nan") &
            (temp["Category"] != "") &
            (temp["Category"].str.lower() != "nan")
        ]

        summary_frames.append(temp)

    if not summary_frames:
        return pd.DataFrame()

    combined = pd.concat(summary_frames, ignore_index=True)

    # Keep only key statuses first
    status_order = ["Open", "On Progress", "Approved"]

    combined["Status"] = combined["Status"].replace({
        "on progress": "On Progress",
        "On progress": "On Progress",
        "OPEN": "Open",
        "APPROVED": "Approved"
    })

    pivot = pd.pivot_table(
        combined,
        index="Status",
        columns="Category",
        values="Category",
        aggfunc="count",
        fill_value=0
    )

    pivot["Total Document"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()

    # Reorder columns
    cols = ["Status", "Total Document"] + [
        c for c in pivot.columns if c not in ["Status", "Total Document"]
    ]
    pivot = pivot[cols]

    # Show Open / On Progress / Approved first
    pivot["sort_order"] = pivot["Status"].apply(
        lambda x: status_order.index(x) if x in status_order else 99
    )
    pivot = pivot.sort_values("sort_order").drop(columns=["sort_order"])

    return pivot


raw_summary = build_raw_status_summary(uploaded_file)

if raw_summary.empty:
    st.info("Upload Excel file with Sheet RFA / RFI to show Raw Data Status Summary.")
else:
    st.dataframe(raw_summary, use_container_width=True, hide_index=True)


# -----------------------------
# Urgent / Overdue Section
# -----------------------------
st.subheader("🚨 Overdue / Urgent Items")

urgent_df = filtered_df[filtered_df["Calculated Status"] == "Overdue"]

if urgent_df.empty:
    st.success("No overdue item found.")
else:
    st.dataframe(
        urgent_df[
            [
                "Document ID",
                "Category",
                "Title / Description",
                "Action By",
                "Due Date",
                "Status Display",
                "Remark"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

# -----------------------------
# Main Table
# -----------------------------
st.subheader("📁 Master Document Register")

display_df = filtered_df.copy()
display_df["Status"] = display_df["Status Display"]

column_config = {
    "Aconex / Link": st.column_config.LinkColumn("Aconex / Link"),
    "Submission Date": st.column_config.DateColumn("Submission Date"),
    "Due Date": st.column_config.DateColumn("Due Date"),
    "Closed Date": st.column_config.DateColumn("Closed Date"),
}

edited_df = st.data_editor(
    display_df[STANDARD_COLUMNS],
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    column_config=column_config,
    disabled=["No"],
)

# -----------------------------
# Downloads
# -----------------------------
st.divider()
st.subheader("⬇️ Download")

download_col1, download_col2 = st.columns(2)

with download_col1:
    template_df = create_template()
    st.download_button(
        label="Download Blank Template",
        data=to_excel_bytes(template_df),
        file_name="Topaz_Tracker_Template_V6_1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with download_col2:
    st.download_button(
        label="Download Current Filtered Data",
        data=to_excel_bytes(filtered_df[STANDARD_COLUMNS]),
        file_name="Topaz_Tracker_Filtered_V6_1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# Footer
# -----------------------------
st.caption("V6.1 Enterprise | Built for Topaz BKK1 ICT Document Tracking")
