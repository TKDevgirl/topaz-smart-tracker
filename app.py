import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd

st.set_page_config(
    page_title="Topaz Smart Document Tracker",
    page_icon="💎",
    layout="wide"
)

tracking_sheets = ["RFA", "RFI"]
takenaka_sheets = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

# =========================
# STYLE
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #f8fbff 0%, #eef4ff 55%, #f8fafc 100%);
    color: #0f172a;
}

.block-container {
    max-width: 1550px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #071226 0%, #0f172a 55%, #111827 100%);
}

[data-testid="stSidebar"] * {
    color: white;
}

.sidebar-logo {
    font-size: 28px;
    font-weight: 900;
    letter-spacing: 1px;
    margin-bottom: 4px;
}

.sidebar-sub {
    color: #c7d2fe;
    font-size: 13px;
    margin-bottom: 24px;
}

.user-card {
    padding: 16px;
    border-radius: 18px;
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.12);
    margin-bottom: 18px;
}

.nav-item {
    padding: 10px 12px;
    border-radius: 12px;
    margin: 6px 0;
    background: rgba(255,255,255,0.04);
}

.hero {
    background: linear-gradient(90deg, #0b1b4d 0%, #172554 52%, #1e3a8a 100%);
    color: white;
    border-radius: 22px;
    padding: 26px 30px;
    box-shadow: 0 16px 35px rgba(15, 23, 42, 0.18);
    margin-bottom: 18px;
}

.hero-grid {
    display: grid;
    grid-template-columns: 1fr auto;
    gap: 20px;
    align-items: center;
}

.hero-title {
    font-size: 38px;
    font-weight: 900;
    margin-bottom: 6px;
}

.hero-sub {
    color: #dbeafe;
    font-size: 15px;
}

.phase-pill {
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.22);
    padding: 10px 16px;
    border-radius: 999px;
    font-weight: 800;
    text-align: center;
}

.info-box {
    padding: 15px 18px;
    border-radius: 16px;
    background: linear-gradient(90deg, #eff6ff, #ffffff);
    border: 1px solid #bfdbfe;
    color: #1d4ed8;
    margin: 16px 0 18px 0;
    font-weight: 650;
}

.admin-box {
    padding: 16px;
    border-radius: 18px;
    background: linear-gradient(90deg, #f0fdf4, #ffffff);
    border: 1px solid #bbf7d0;
    color: #166534;
    margin: 16px 0;
    font-weight: 650;
}

.kpi-card {
    border-radius: 22px;
    background: rgba(255,255,255,0.95);
    box-shadow: 0 12px 28px rgba(15, 23, 42, 0.08);
    border: 1px solid #e2e8f0;
    padding: 22px;
    min-height: 158px;
    position: relative;
    overflow: hidden;
}

.kpi-card::after {
    content: "";
    position: absolute;
    left: 0;
    right: 0;
    bottom: 0;
    height: 5px;
    background: var(--accent);
}

.kpi-icon {
    height: 48px;
    width: 48px;
    border-radius: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--soft);
    color: var(--accent);
    font-size: 24px;
    margin-bottom: 10px;
}

.kpi-title {
    color: #64748b;
    font-size: 13px;
    font-weight: 800;
}

.kpi-value {
    font-size: 36px;
    font-weight: 900;
    color: #0f172a;
    line-height: 1.1;
    margin: 6px 0;
}

.kpi-sub {
    color: #64748b;
    font-size: 13px;
}

.panel {
    background: rgba(255,255,255,0.95);
    border: 1px solid #e2e8f0;
    border-radius: 24px;
    box-shadow: 0 12px 28px rgba(15, 23, 42, 0.07);
    padding: 22px;
    margin-bottom: 18px;
}

.panel-title {
    font-size: 21px;
    font-weight: 900;
    color: #0f172a;
    margin-bottom: 14px;
}

.project-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0 6px;
    font-size: 14px;
}

.project-table td {
    padding: 9px 10px;
    background: #f8fafc;
    border-top: 1px solid #e2e8f0;
    border-bottom: 1px solid #e2e8f0;
}

.project-table td:first-child {
    font-weight: 900;
    color: #1e3a8a;
    border-left: 1px solid #e2e8f0;
    border-radius: 10px 0 0 10px;
    width: 40%;
}

.project-table td:last-child {
    border-right: 1px solid #e2e8f0;
    border-radius: 0 10px 10px 0;
}

.health-box {
    background: linear-gradient(90deg, #fee2e2, #fff7ed);
    border: 1px solid #fca5a5;
    color: #991b1b;
    padding: 14px;
    border-radius: 16px;
    font-weight: 800;
    margin-top: 12px;
}

.quick-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px;
    border-radius: 14px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    margin-bottom: 8px;
}

.count-pill {
    padding: 4px 10px;
    border-radius: 999px;
    font-weight: 900;
    background: #dbeafe;
    color: #1d4ed8;
}
</style>
""", unsafe_allow_html=True)

# =========================
# LOGIN
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "role" not in st.session_state:
    st.session_state.role = "viewer"
if "username" not in st.session_state:
    st.session_state.username = ""

with st.sidebar:
    st.markdown('<div class="sidebar-logo">💎 TOPAZ</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Smart Document Tracker v3.0</div>', unsafe_allow_html=True)

    if not st.session_state.logged_in:
        username = st.text_input("Username")
        if st.button("Login", use_container_width=True):
            st.session_state.username = username.strip() or "viewer"
            st.session_state.role = "admin" if username.strip().lower() == "pavinee" else "viewer"
            st.session_state.logged_in = True
            st.rerun()
    else:
        st.markdown(
            f"""
            <div class="user-card">
                <b>👤 User</b><br>{st.session_state.username}<br><br>
                <b>🔑 Role</b><br>{st.session_state.role.title()}
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.button("Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.role = "viewer"
            st.session_state.username = ""
            st.rerun()

    st.divider()
    st.markdown('<div class="nav-item">🏠 Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">⬇ Download Report</div>', unsafe_allow_html=True)

# =========================
# FUNCTIONS
# =========================
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
        "Takenaka Sheet", "Takenaka Doc No", "Takenaka Status 1",
        "Takenaka Status 2", "Takenaka Status 3", "Action", "Checked Time"
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

            # Tracking structure: E = Status, F = Info
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
                    "", "", "", "", "", action, checked_time
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

# =========================
# SESSION STATE
# =========================
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

# =========================
# HERO
# =========================
st.markdown(
    """
    <div class="hero">
        <div class="hero-grid">
            <div>
                <div class="hero-title">💎 Topaz Smart Document Tracker</div>
                <div class="hero-sub">TOPAZ BKK1 | ICT Project Dashboard | Phase 1A — NTP1</div>
            </div>
            <div class="phase-pill">Data control dashboard</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================
# UPLOAD
# =========================
if st.session_state.role == "admin":
    st.markdown('<div class="admin-box">👩‍💼 Admin mode: Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

    up1, up2, up3 = st.columns([1, 1, 0.8])
    with up1:
        tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
    with up2:
        takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])
    with up3:
        st.write("")
        st.write("")
        generate_btn = st.button("🚀 Generate Dashboard", type="primary", use_container_width=True)

    if tracking_file and takenaka_file and generate_btn:
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
    st.markdown('<div class="info-box">ℹ️ Viewer mode: only Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

# =========================
# DASHBOARD
# =========================
if st.session_state.result_df is not None:
    df = st.session_state.result_df
    action_counts = st.session_state.action_counts

    open_process = action_counts.get("OPEN & ON PROCESS", 0)
    update_closed = action_counts.get("UPDATE TRACKING TO CLOSED", 0)
    returned = action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)
    overdue = action_counts.get("OVERDUE / FOLLOW UP", 0)
    not_found = action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)
    check = action_counts.get("CHECK", 0)

    st.success("Dashboard generated successfully ✅")

    k1, k2, k3, k4, k5 = st.columns(5)
    cards = [
        ("📁", "Total Documents", st.session_state.total_docs, "All documents", "#7c3aed", "#ede9fe"),
        ("🕒", "Open / On Progress", st.session_state.focus_docs, "Require review", "#2563eb", "#dbeafe"),
        ("▶", "Open & On Process", open_process, "Waiting review", "#16a34a", "#dcfce7"),
        ("🔄", "Need Update", update_closed, "Update tracking", "#f97316", "#ffedd5"),
        ("⚠", "Overdue", overdue, "Follow up", "#ef4444", "#fee2e2"),
    ]
    for col, (icon, title, value, sub, accent, soft) in zip([k1, k2, k3, k4, k5], cards):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card" style="--accent:{accent}; --soft:{soft};">
                    <div class="kpi-icon">{icon}</div>
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.write("")

    chart_col, info_col = st.columns([2, 1])

    with chart_col:
        st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)
        summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])
        if not summary_df.empty:
            st.bar_chart(summary_df.set_index("Action"))
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with info_col:
        st.markdown(
            """
            <div class="panel">
                <div class="panel-title">🏗 Project Information</div>
                <table class="project-table">
                    <tr><td>Discipline</td><td>ICT</td></tr>
                    <tr><td>Consultant</td><td>NV5</td></tr>
                    <tr><td>Contractor</td><td>ByteBridge</td></tr>
                    <tr><td>Current Phase</td><td>Phase 1A (NTP1)</td></tr>
                    <tr><td>Current Focus</td><td>Fiber Tray / ODF / PDU</td></tr>
                    <tr><td>Project</td><td>TOPAZ BKK1</td></tr>
                    <tr><td>Client</td><td>TTI</td></tr>
                </table>
            """,
            unsafe_allow_html=True
        )

        urgent = returned + overdue + not_found
        st.markdown(
            f"""
            <div class="health-box">
                🚨 Project Health: Attention<br>
                {urgent} items require follow-up action.
            </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    quick_col, progress_col = st.columns([1, 2])

    with quick_col:
        st.markdown('<div class="panel"><div class="panel-title">⚡ Quick Action</div>', unsafe_allow_html=True)
        quick_items = [
            ("Returned by NV5", returned),
            ("Need update closed", update_closed),
            ("Overdue follow up", overdue),
            ("Not found in Takenaka", not_found),
            ("Check manually", check),
        ]
        for label, count in quick_items:
            st.markdown(
                f'<div class="quick-row"><span>{label}</span><span class="count-pill">{count}</span></div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

    with progress_col:
        st.markdown('<div class="panel"><div class="panel-title">📡 ICT Progress Status</div>', unsafe_allow_html=True)
        progress_data = pd.DataFrame([
            {"Topic": "Fiber Tray", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing", "Progress": "75%"},
            {"Topic": "ODF", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing", "Progress": "65%"},
            {"Topic": "PDU", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing", "Progress": "60%"},
            {"Topic": "Rack", "Status": "🔵 Learning", "Meaning / Action": "Reviewing drawings, preparing site", "Progress": "40%"},
            {"Topic": "Backbone", "Status": "🔵 Learning", "Meaning / Action": "Reviewing drawings, preparing site", "Progress": "40%"},
            {"Topic": "Deviation Review", "Status": "🟢 Working", "Meaning / Action": "Installation ongoing", "Progress": "70%"},
            {"Topic": "Commissioning", "Status": "⚪ Not Start", "Meaning / Action": "Pending system completion", "Progress": "0%"},
        ])
        st.dataframe(progress_data, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)

    f1, f2, f3 = st.columns([1, 2, 0.8])
    with f1:
        selected_action = st.selectbox("Filter by Action", ["All"] + sorted(df["Action"].dropna().unique().tolist()))
    with f2:
        search = st.text_input("Search Document No / Document Name")
    with f3:
        st.write("")
        st.write("")
        st.download_button(
            label="⬇️ Export Excel",
            data=st.session_state.report,
            file_name="Open_On_Process_Compare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    filtered_df = df.copy()
    if selected_action != "All":
        filtered_df = filtered_df[filtered_df["Action"] == selected_action]

    if search:
        filtered_df = filtered_df[
            filtered_df.astype(str).apply(
                lambda x: x.str.contains(search, case=False, na=False)
            ).any(axis=1)
        ]

    st.dataframe(filtered_df, use_container_width=True, height=430)
    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("Please upload both Excel files and click Generate Dashboard.")
