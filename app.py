import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd

st.set_page_config(
    page_title="TOPAZ BKK1 ICT Dashboard",
    page_icon="💎",
    layout="wide"
)

tracking_sheets = ["RFA", "RFI"]
takenaka_sheets = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

st.markdown("""
<style>
.stApp { background: #050816; color: white; }
[data-testid="stSidebar"] { background: #061124; }
[data-testid="stSidebar"] * { color: white; }
.block-container { padding-top: 1rem; max-width: 1500px; }

.header {
    background: linear-gradient(90deg,#0b1b4d,#111a5c);
    border: 1px solid #243b82;
    border-radius: 8px;
    padding: 14px;
    display: grid;
    grid-template-columns: 1fr 2fr 1fr;
    text-align: center;
    font-weight: 900;
    margin-bottom: 18px;
}

.title {
    font-size: 32px;
    font-weight: 900;
    margin-bottom: 6px;
}

.subtitle {
    color: #a5b4fc;
    margin-bottom: 18px;
}

.info {
    background: #0b1b4d;
    border: 1px solid #3b82f6;
    padding: 14px;
    border-radius: 8px;
    margin-bottom: 18px;
}

.kpi {
    background: linear-gradient(180deg,#0f1b3d,#0b132b);
    border: 1px solid #263b75;
    border-radius: 14px;
    padding: 20px;
    box-shadow: 0 10px 25px rgba(0,0,0,.25);
    min-height: 145px;
}

.kpi-title { font-size: 14px; color: #bfdbfe; font-weight: 800; }
.kpi-value { font-size: 38px; font-weight: 900; margin: 8px 0; }
.kpi-sub { color: #94a3b8; font-size: 13px; }

.panel {
    background: #0b132b;
    border: 1px solid #263b75;
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 18px;
}

.panel-title {
    font-size: 20px;
    font-weight: 900;
    margin-bottom: 12px;
}

.health {
    background: linear-gradient(90deg,#7f1d1d,#450a0a);
    border: 1px solid #ef4444;
    border-radius: 10px;
    padding: 14px;
    font-weight: 800;
}

.project-table {
    width: 100%;
    border-collapse: collapse;
}
.project-table td {
    border: 1px solid #263b75;
    padding: 7px;
}
.project-table td:first-child {
    font-weight: 900;
    color: #dbeafe;
}
</style>
""", unsafe_allow_html=True)


# ================= LOGIN =================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "role" not in st.session_state:
    st.session_state.role = "viewer"
if "username" not in st.session_state:
    st.session_state.username = ""

with st.sidebar:
    st.markdown("## 💎 TOPAZ")
    st.markdown("### Smart Tracker")
    st.divider()

    if not st.session_state.logged_in:
        username = st.text_input("Username")
        if st.button("Login"):
            st.session_state.username = username.strip()
            st.session_state.role = "admin" if username.strip().lower() == "pavinee" else "viewer"
            st.session_state.logged_in = True
            st.rerun()
    else:
        st.markdown(f"👤 *User:* {st.session_state.username}")
        st.markdown(f"🔑 *Role:* {st.session_state.role.title()}")
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
    st.caption("Topaz Smart Document Tracker v3.0")


# ================= FUNCTIONS =================
def base_doc_no(doc_no):
    doc_no = str(doc_no).strip()
    parts = doc_no.split("-")
    if parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])
    return doc_no


def read_takenaka(takenaka_file):
    wb = load_workbook(takenaka_file, data_only=True)
    data = {}

    for sheet in takenaka_sheets:
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]

        for row in range(11, ws.max_row + 1):
            doc_no = ws[f"E{row}"].value
            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            data[base_doc_no(doc_no)] = {
                "Takenaka Sheet": sheet,
                "Takenaka Doc No": doc_no,
                "Takenaka Status 1": ws[f"AA{row}"].value,
                "Takenaka Status 2": ws[f"AB{row}"].value,
                "Takenaka Status 3": ws[f"AC{row}"].value,
            }

    return data


def generate_report(tracking_file, takenaka_file):
    takenaka_map = read_takenaka(takenaka_file)
    wb = load_workbook(tracking_file)

    report_sheet = "Open_On_Process_Compare"
    if report_sheet in wb.sheetnames:
        del wb[report_sheet]

    report_ws = wb.create_sheet(report_sheet)

    headers = [
        "Tracking Sheet", "Document No", "Document Name", "Tracking Status", "Info",
        "Takenaka Sheet", "Takenaka Doc No",
        "Takenaka Status 1", "Takenaka Status 2", "Takenaka Status 3",
        "Action", "Checked Time"
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
    focus_docs = 0

    for sheet in tracking_sheets:
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]

        for row in range(2, ws.max_row + 1):
            doc_no = ws[f"B{row}"].value
            doc_name = ws[f"D{row}"].value

            # Tracking file columns:
            # E = Status, F = Info
            tracking_status = ws[f"E{row}"].value
            info = ws[f"F{row}"].value

            if not doc_no:
                continue

            total_docs += 1

            status_text = str(tracking_status or "").strip().upper()
            info_text = str(info or "").strip().upper()

            if (
                "OPEN" not in status_text
                and "ON PROGRESS" not in status_text
                and "ON PROCESS" not in status_text
                and "ON PROGRESS" not in info_text
                and "ON PROCESS" not in info_text
            ):
                continue

            focus_docs += 1
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
                    sheet, doc_no, doc_name, tracking_status, info,
                    src["Takenaka Sheet"], src["Takenaka Doc No"],
                    src["Takenaka Status 1"], src["Takenaka Status 2"], src["Takenaka Status 3"],
                    action, checked_time
                ]
            else:
                action = "NOT FOUND IN TAKENAKA SOURCE"
                fill = red
                new_row = [
                    sheet, doc_no, doc_name, tracking_status, info,
                    "", "", "", "", "",
                    action, checked_time
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

    return output, total_docs, focus_docs, rows


# ================= SESSION =================
defaults = {
    "result_df": None,
    "report": None,
    "total_docs": 0,
    "focus_docs": 0,
    "action_counts": {}
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ================= HEADER =================
st.markdown("""
<div class="header">
    <div>TOPAZ BKK1</div>
    <div>ICT PROJECT DASHBOARD</div>
    <div>PHASE 1A — NTP1</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="title">Project Engineer Dashboard</div>', unsafe_allow_html=True)


# ================= UPLOAD =================
if st.session_state.role == "admin":
    st.markdown('<div class="info">👩‍💼 Admin mode: Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

    tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
    takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

    if tracking_file and takenaka_file:
        if st.button("🚀 Generate Dashboard", type="primary"):
            with st.spinner("Reading files and generating dashboard..."):
                report, total_docs, focus_docs, rows = generate_report(tracking_file, takenaka_file)

            df = pd.DataFrame(rows)
            action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}

            st.session_state.result_df = df
            st.session_state.report = report
            st.session_state.total_docs = total_docs
            st.session_state.focus_docs = focus_docs
            st.session_state.action_counts = action_counts
            st.rerun()
else:
    st.markdown('<div class="info">ℹ️ Viewer mode: only Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)


# ================= DASHBOARD =================
if st.session_state.result_df is not None:
    df = st.session_state.result_df
    action_counts = st.session_state.action_counts

    open_process = action_counts.get("OPEN & ON PROCESS", 0)
    update_closed = action_counts.get("UPDATE TRACKING TO CLOSED", 0)
    returned = action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)
    overdue = action_counts.get("OVERDUE / FOLLOW UP", 0)
    not_found = action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)

    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-title">📁 Open MDR</div>
            <div class="kpi-value">{st.session_state.focus_docs}</div>
            <div class="kpi-sub">Tracking items awaiting review</div>
        </div>
        """, unsafe_allow_html=True)

    with k2:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-title">📄 Pending Review</div>
            <div class="kpi-value">{open_process}</div>
            <div class="kpi-sub">Open and on process</div>
        </div>
        """, unsafe_allow_html=True)

    with k3:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-title">🔴 Returned</div>
            <div class="kpi-value">{returned}</div>
            <div class="kpi-sub">Revise and resubmit</div>
        </div>
        """, unsafe_allow_html=True)

    with k4:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-title">⚠ Overdue</div>
            <div class="kpi-value">{overdue}</div>
            <div class="kpi-sub">Follow-up required</div>
        </div>
        """, unsafe_allow_html=True)

    with k5:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-title">✅ Approved / Close</div>
            <div class="kpi-value">{update_closed}</div>
            <div class="kpi-sub">Update tracking to closed</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    left, right = st.columns([2.2, 1])

    with left:
        st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)
        summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])
        if not summary_df.empty:
            st.bar_chart(summary_df.set_index("Action"))
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown("""
        <div class="panel">
        <div class="panel-title">Project Information</div>
        <table class="project-table">
            <tr><td>Discipline</td><td>ICT</td></tr>
            <tr><td>Consultant</td><td>NV5</td></tr>
            <tr><td>Contractor</td><td>ByteBridge</td></tr>
            <tr><td>Current Phase</td><td>Phase 1A (NTP1)</td></tr>
            <tr><td>Current Focus</td><td>Fiber Tray / ODF / PDU</td></tr>
            <tr><td>Project</td><td>TOPAZ BKK1</td></tr>
            <tr><td>Client</td><td>TTI</td></tr>
        </table>
        """, unsafe_allow_html=True)

        urgent = returned + overdue + not_found
        st.markdown(f"""
        <div class="health">
            🚨 Project Health: Attention<br>
            {urgent} items require follow-up action.
        </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">ICT Progress Status</div>', unsafe_allow_html=True)
    progress_data = pd.DataFrame([
        {"Topic": "Fiber Tray", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing"},
        {"Topic": "ODF", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing"},
        {"Topic": "PDU", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing"},
        {"Topic": "Rack", "Status": "🔵 Learning", "Meaning / Action": "Reviewing drawings, preparing site"},
        {"Topic": "Backbone", "Status": "🔵 Learning", "Meaning / Action": "Reviewing drawings, preparing site"},
        {"Topic": "Deviation Review", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing"},
        {"Topic": "Commissioning", "Status": "⚪ Not Start", "Meaning / Action": "Pending system completion"},
    ])
    st.dataframe(progress_data, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)

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

    st.dataframe(filtered_df, use_container_width=True, height=420)

    st.download_button(
        label="⬇️ Download Excel Report",
        data=st.session_state.report,
        file_name="Open_On_Process_Compare.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("Please upload both Excel files and click Generate Dashboard.")
