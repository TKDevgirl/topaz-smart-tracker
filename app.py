import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="Topaz Smart Document Tracker",page_icon="💎",layout="wide")

tracking_sheets = ["RFA", "RFI"]takenaka_sheets = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

=========================

STYLE

=========================

st.markdown("""

<style>
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #111827 100%);
}
[data-testid="stSidebar"] * {
    color: white;
}
.main-title {
    font-size: 40px;
    font-weight: 800;
    color: #172554;
}
.sub-title {
    color: #64748b;
    font-size: 15px;
}
.info-box {
    padding: 18px;
    border-radius: 14px;
    background: linear-gradient(90deg, #eff6ff, #f8fafc);
    border: 1px solid #bfdbfe;
    margin: 20px 0;
}
.kpi-card {
    padding: 24px;
    border-radius: 18px;
    background: white;
    box-shadow: 0 6px 20px rgba(15, 23, 42, 0.08);
    border: 1px solid #e5e7eb;
}
.kpi-title {
    font-size: 14px;
    color: #64748b;
}
.kpi-value {
    font-size: 34px;
    font-weight: 800;
    color: #111827;
}
.badge {
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
}
</style>

""", unsafe_allow_html=True)

=========================

LOGIN

=========================

if "logged_in" not in st.session_state:st.session_state.logged_in = Falseif "role" not in st.session_state:st.session_state.role = "viewer"if "username" not in st.session_state:st.session_state.username = ""

with st.sidebar:st.markdown("## 💎 TOPAZ")st.markdown("### Smart Tracker")st.divider()

if not st.session_state.logged_in:
    username = st.text_input("Username")

    if st.button("Login"):
        st.session_state.username = username.strip()

        if username.strip().lower() == "pavinee":
            st.session_state.role = "admin"
        else:
            st.session_state.role = "viewer"

        st.session_state.logged_in = True
        st.rerun()
else:
    st.markdown(f"👤 **User:** {st.session_state.username}")
    st.markdown(f"🔑 **Role:** {st.session_state.role.title()}")

    if st.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.role = "viewer"
        st.session_state.username = ""
        st.rerun()

st.divider()
st.markdown("🏠 Dashboard")
st.markdown("📋 Documents")
st.markdown("📊 Action Summary")
st.markdown("⬇ Download Report")
st.divider()
st.caption("Topaz Smart Document Tracker v2.0")

=========================

FUNCTIONS

=========================

def base_doc_no(doc_no):doc_no = str(doc_no).strip()parts = doc_no.split("-")if parts[-1].isdigit() and len(parts[-1]) == 2:return "-".join(parts[:-1])return doc_no

def read_takenaka(takenaka_file):wb = load_workbook(takenaka_file, data_only=True)data = {}

for sheet in takenaka_sheets:
    if sheet not in wb.sheetnames:
        continue

    ws = wb[sheet]

    for row in range(11, ws.max_row + 1):
        doc_no = ws[f"E{row}"].value

        if not doc_no or "DETH-NSC" not in str(doc_no):
            continue

        key = base_doc_no(doc_no)

        data[key] = {
            "Takenaka Sheet": sheet,
            "Takenaka Doc No": doc_no,
            "Takenaka Status 1": ws[f"AA{row}"].value,
            "Takenaka Status 2": ws[f"AB{row}"].value,
            "Takenaka Status 3": ws[f"AC{row}"].value,
        }

return data

def generate_report(tracking_file, takenaka_file):takenaka_map = read_takenaka(takenaka_file)

wb = load_workbook(tracking_file)
report_sheet = "Open_On_Process_Compare"

if report_sheet in wb.sheetnames:
    del wb[report_sheet]

report_ws = wb.create_sheet(report_sheet)

headers = [
    "Tracking Sheet",
    "Document No",
    "Document Name",
    "Tracking Status",
    "Info",
    "Takenaka Sheet",
    "Takenaka Doc No",
    "Takenaka Status 1",
    "Takenaka Status 2",
    "Takenaka Status 3",
    "Action",
    "Checked Time",
]

report_ws.append(headers)

for cell in report_ws[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

green = PatternFill(fill_type="solid", fgColor="C6EFCE")
yellow = PatternFill(fill_type="solid", fgColor="FFEB9C")
red = PatternFill(fill_type="solid", fgColor="FFC7CE")
blue = PatternFill(fill_type="solid", fgColor="BDD7EE")

rows = []
total_docs = 0
open_docs = 0

for sheet in tracking_sheets:
    if sheet not in wb.sheetnames:
        continue

    ws = wb[sheet]

    for row in range(2, ws.max_row + 1):
        doc_no = ws[f"B{row}"].value
        doc_name = ws[f"D{row}"].value
       tracking_status = ws[f"E{row}"].value
       info = ws[f"F{row}"].value

        if not doc_no:
            continue

        total_docs += 1

        tracking_status_text = str(tracking_status or "").strip().upper()
        info_text = str(info or "").strip().upper()

        if (
            "OPEN" not in tracking_status_text
            and "ON PROGRESS" not in tracking_status_text
            and "ON PROCESS" not in tracking_status_text
            and "ON PROGRESS" not in info_text
            and "ON PROCESS" not in info_text
        ):
            continue

        open_docs += 1
        key = base_doc_no(doc_no)
        checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

        if key in takenaka_map:
            src = takenaka_map[key]

            s1 = str(src["Takenaka Status 1"] or "").strip().upper()
            s2 = str(src["Takenaka Status 2"] or "").strip().upper()
            s3 = str(src["Takenaka Status 3"] or "").strip().upper()

            if s1 == "CLOSED":
                action = "UPDATE TRACKING TO CLOSED"
                fill = green
            elif s1 == "RETURNED":
                action = "RETURNED BY NV5 / NEED RESUBMIT"
                fill = red
            elif s3 == "OVERDUE":
                action = "OVERDUE / FOLLOW UP"
                fill = red
            elif s1 == "OPEN" and s2 == "ON PROCESS":
                action = "OPEN & ON PROCESS"
                fill = yellow
            elif s1 == "OPEN":
                action = "OPEN"
                fill = blue
            else:
                action = "CHECK"
                fill = blue

            new_row = [
                sheet,
                doc_no,
                doc_name,
                tracking_status,
                info,
                src["Takenaka Sheet"],
                src["Takenaka Doc No"],
                src["Takenaka Status 1"],
                src["Takenaka Status 2"],
                src["Takenaka Status 3"],
                action,
                checked_time,
            ]

        else:
            action = "NOT FOUND IN TAKENAKA SOURCE"
            fill = red

            new_row = [
                sheet,
                doc_no,
                doc_name,
                tracking_status,
                info,
                "",
                "",
                "",
                "",
                "",
                action,
                checked_time,
            ]

        report_ws.append(new_row)
        report_ws[f"K{report_ws.max_row}"].fill = fill

        rows.append({
            "Tracking Sheet": new_row[0],
            "Document No": new_row[1],
            "Document Name": new_row[2],
            "Tracking Status": new_row[3],
            "Info": new_row[4],
            "Takenaka Status 1": new_row[7],
            "Takenaka Status 2": new_row[8],
            "Takenaka Status 3": new_row[9],
            "Action": new_row[10],
            "Checked Time": new_row[11],
        })

for col in report_ws.columns:
    max_len = 0
    col_letter = col[0].column_letter
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    report_ws.column_dimensions[col_letter].width = min(max_len + 2, 55)

output = BytesIO()
wb.save(output)
output.seek(0)

return output, total_docs, open_docs, rows

=========================

SESSION STATE

=========================

if "result_df" not in st.session_state:st.session_state.result_df = Nonest.session_state.report = Nonest.session_state.total_docs = 0st.session_state.open_docs = 0st.session_state.action_counts = {}

=========================

HEADER

=========================

col_title, col_time = st.columns([3, 1])

with col_title:st.markdown('<div class="main-title">📄 Topaz Smart Document Tracker</div>', unsafe_allow_html=True)st.markdown('<div class="sub-title">Web dashboard for OPEN / ON PROGRESS documents compared with Takenaka status.</div>', unsafe_allow_html=True)

with col_time:st.info("Data is up to date")

=========================

UPLOAD AREA

=========================

if st.session_state.role == "admin":st.markdown('<div class="info-box">👩‍💼 Admin mode: Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

if tracking_file and takenaka_file:
    if st.button("🚀 Generate Dashboard", type="primary"):
        with st.spinner("Reading files and generating dashboard..."):
            report, total_docs, open_docs, rows = generate_report(
                tracking_file,
                takenaka_file
            )

        df = pd.DataFrame(rows)
        action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}

        st.session_state.result_df = df
        st.session_state.report = report
        st.session_state.total_docs = total_docs
        st.session_state.open_docs = open_docs
        st.session_state.action_counts = action_counts

else:st.markdown('<div class="info-box">ℹ️ Viewer mode: only Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

=========================

DASHBOARD

=========================

if st.session_state.result_df is not None:df = st.session_state.result_dfaction_counts = st.session_state.action_counts

st.success("Dashboard generated successfully ✅")

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">📁 Total Documents</div>
        <div class="kpi-value">{st.session_state.total_docs}</div>
        <div class="kpi-title">All documents</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">🕒 Open / On Progress</div>
        <div class="kpi-value">{st.session_state.open_docs}</div>
        <div class="kpi-title">Documents</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">▶ Open & On Process</div>
        <div class="kpi-value">{action_counts.get("OPEN & ON PROCESS", 0)}</div>
        <div class="kpi-title">Waiting review</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">🔄 Need Update</div>
        <div class="kpi-value">{action_counts.get("UPDATE TRACKING TO CLOSED", 0)}</div>
        <div class="kpi-title">Update tracking</div>
    </div>
    """, unsafe_allow_html=True)

with c5:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">⚠ Overdue</div>
        <div class="kpi-value">{action_counts.get("OVERDUE / FOLLOW UP", 0)}</div>
        <div class="kpi-title">Follow up</div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

chart_col, table_col = st.columns([1, 1])

with chart_col:
    st.subheader("📊 Action Summary")
    summary_df = pd.DataFrame(
        [{"Action": k, "Count": v} for k, v in action_counts.items()]
    )

    if not summary_df.empty:
        st.bar_chart(summary_df.set_index("Action"))
        st.dataframe(summary_df, use_container_width=True)

with table_col:
    st.subheader("🧭 Quick Action")
    st.write("Returned by NV5:", action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0))
    st.write("Need update closed:", action_counts.get("UPDATE TRACKING TO CLOSED", 0))
    st.write("Overdue follow up:", action_counts.get("OVERDUE / FOLLOW UP", 0))
    st.write("Not found in Takenaka:", action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0))

st.divider()

st.subheader("📋 Document Action List")

selected_action = st.selectbox(
    "Filter by Action",
    ["All"] + sorted(df["Action"].dropna().unique().tolist())
)

search = st.text_input("Search Document No / Document Name")

filtered_df = df.copy()

if selected_action != "All":
    filtered_df = filtered_df[filtered_df["Action"] == selected_action]

if search:
    filtered_df = filtered_df[
        filtered_df.astype(str).apply(
            lambda x: x.str.contains(search, case=False, na=False)
        ).any(axis=1)
    ]

st.dataframe(filtered_df, use_container_width=True, height=480)

st.download_button(
    label="⬇️ Download Excel Report",
    data=st.session_state.report,
    file_name="Open_On_Process_Compare.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

else:st.info("Please upload both Excel files and click Generate Dashboard.")
