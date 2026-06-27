import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd
import altair as alt

st.set_page_config(page_title="Topaz Smart Document Tracker V4 Clean", page_icon="💎", layout="wide")

TRACKING_SHEETS = ["RFA", "RFI"]
TAKENAKA_SHEETS = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

# =============================
# STYLE
# =============================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f4f7fb; color: #0f172a; }
.block-container { max-width: 1580px; padding-top: 1.2rem; padding-bottom: 3rem; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #061124 0%, #0f172a 72%, #111827 100%);
}
[data-testid="stSidebar"] * { color: white; }

.logo { font-size: 30px; font-weight: 900; margin-bottom: 2px; }
.logo-sub { color:#c7d2fe; font-size: 13px; margin-bottom: 22px; }

.user-card {
    padding: 16px;
    border-radius: 18px;
    background: rgba(255,255,255,.08);
    border: 1px solid rgba(255,255,255,.14);
    margin-bottom: 18px;
}

.nav-item {
    padding: 11px 13px;
    border-radius: 13px;
    margin: 7px 0;
    background: rgba(255,255,255,.05);
    font-weight: 750;
}

.nav-active {
    padding: 11px 13px;
    border-radius: 13px;
    margin: 7px 0;
    background: linear-gradient(90deg,#4f46e5,#7c3aed);
    font-weight: 850;
}

.hero {
    background: radial-gradient(circle at top right, rgba(96,165,250,.42), transparent 28%),
                linear-gradient(90deg,#071226 0%, #0b1b4d 48%, #172554 100%);
    color: white;
    border-radius: 26px;
    padding: 28px 32px;
    box-shadow: 0 20px 40px rgba(15,23,42,.20);
    margin-bottom: 18px;
}

.hero-grid {
    display:grid;
    grid-template-columns:1fr auto;
    gap:20px;
    align-items:center;
}

.hero-title {
    font-size: 38px;
    font-weight: 950;
    margin-bottom: 7px;
}

.hero-sub {
    color:#dbeafe;
    font-size:15px;
}

.pill {
    background: rgba(255,255,255,.12);
    border:1px solid rgba(255,255,255,.22);
    border-radius: 999px;
    padding: 10px 16px;
    font-weight: 850;
}

.notice {
    padding: 15px 18px;
    border-radius: 18px;
    margin: 16px 0 18px 0;
    font-weight: 700;
    background: linear-gradient(90deg,#eff6ff,#ffffff);
    border:1px solid #bfdbfe;
    color:#1d4ed8;
}

.admin {
    padding: 15px 18px;
    border-radius: 18px;
    margin: 16px 0 18px 0;
    font-weight: 700;
    background: linear-gradient(90deg,#ecfdf5,#ffffff);
    border:1px solid #bbf7d0;
    color:#166534;
}

.kpi-card {
    background: rgba(255,255,255,.96);
    border: 1px solid #e2e8f0;
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 14px 32px rgba(15,23,42,.08);
    min-height: 160px;
    position: relative;
    overflow:hidden;
    transition: .18s ease;
}

.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 18px 38px rgba(15,23,42,.12);
}

.kpi-card:after {
    content:"";
    position:absolute;
    left:0;
    right:0;
    bottom:0;
    height:5px;
    background:var(--accent);
}

.kpi-icon {
    width: 50px;
    height: 50px;
    border-radius: 17px;
    display:flex;
    align-items:center;
    justify-content:center;
    background: var(--soft);
    color: var(--accent);
    font-size:24px;
    margin-bottom:10px;
}

.kpi-title {
    color:#64748b;
    font-size:13px;
    font-weight:850;
}

.kpi-value {
    font-size: 38px;
    font-weight: 950;
    color:#0f172a;
    line-height:1.05;
    margin:7px 0;
}

.kpi-sub {
    color:#64748b;
    font-size:13px;
}

.panel {
    background: rgba(255,255,255,.97);
    border:1px solid #e2e8f0;
    border-radius: 24px;
    box-shadow: 0 14px 32px rgba(15,23,42,.07);
    padding: 22px;
    margin-bottom: 18px;
}

.panel-title {
    font-size: 21px;
    font-weight: 950;
    color:#0f172a;
    margin-bottom: 14px;
}

.quick-row {
    display:flex;
    justify-content:space-between;
    align-items:center;
    padding:12px;
    border-radius:14px;
    background:#f8fafc;
    border:1px solid #e2e8f0;
    margin-bottom:8px;
}

.count-pill {
    padding:4px 10px;
    border-radius:999px;
    font-weight:900;
    background:#dbeafe;
    color:#1d4ed8;
}

.footer-note {
    color:#64748b;
    font-size:12px;
    margin-top:6px;
}
</style>
""", unsafe_allow_html=True)

# =============================
# LOGIN
# =============================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "role" not in st.session_state:
    st.session_state.role = "viewer"
if "username" not in st.session_state:
    st.session_state.username = ""

with st.sidebar:
    st.markdown('<div class="logo">💎 TOPAZ</div>', unsafe_allow_html=True)
    st.markdown('<div class="logo-sub">Smart Document Tracker V4 Clean</div>', unsafe_allow_html=True)

    if not st.session_state.logged_in:
        username = st.text_input("Username")
        if st.button("Login", use_container_width=True):
            st.session_state.username = username.strip() or "viewer"
            st.session_state.role = "admin" if username.strip().lower() == "pavinee" else "viewer"
            st.session_state.logged_in = True
            st.rerun()
    else:
        st.markdown(f"""
        <div class="user-card">
            <b>👤 User</b><br>{st.session_state.username}<br><br>
            <b>🔑 Role</b><br>{st.session_state.role.title()}
        </div>
        """, unsafe_allow_html=True)

        if st.button("Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.role = "viewer"
            st.session_state.username = ""
            st.rerun()

    st.divider()
    st.markdown('<div class="nav-active">🏠 Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">⬇ Export Report</div>', unsafe_allow_html=True)
    st.divider()
    st.caption("© 2026 Topaz Smart Tracker")

# =============================
# DATA FUNCTIONS
# =============================
def norm_text(value):
    return str(value or "").strip()

def norm_upper(value):
    return norm_text(value).upper()

def base_doc_no(doc_no):
    doc_no = norm_text(doc_no)
    parts = doc_no.split("-")
    if parts and parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])
    return doc_no

def normalize_header(value):
    return (
        str(value or "")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .upper()
    )

def find_header_row_and_columns(ws):
    for row in range(1, 25):
        headers = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value
            if value:
                headers[normalize_header(value)] = col

        if "STATUS" in headers and "INFO" in headers:
            return row, headers

    return 1, {
        "DOCUMENT NO": 2,
        "DOCUMENT NAME": 4,
        "STATUS": 5,
        "INFO": 6,
    }

def get_col(headers, possible_names):
    for name in possible_names:
        key = normalize_header(name)
        if key in headers:
            return headers[key]

    for key, col in headers.items():
        for name in possible_names:
            name_key = normalize_header(name)
            if name_key in key or key in name_key:
                return col

    return None

def normalize_action(action):
    a = norm_upper(action)
    if "RETURNED" in a:
        return "RETURNED BY NV5 / NEED RESUBMIT"
    if "OVERDUE" in a:
        return "OVERDUE / FOLLOW UP"
    if "UPDATE" in a or "CLOSED" in a:
        return "UPDATE TRACKING TO CLOSED"
    if "OPEN" in a and ("PROCESS" in a or "PROGRESS" in a):
        return "OPEN & ON PROCESS"
    if "NOT FOUND" in a:
        return "NOT FOUND IN TAKENAKA SOURCE"
    if "OPEN" in a:
        return "OPEN"
    return "CHECK"

def read_takenaka(takenaka_file):
    wb = load_workbook(takenaka_file, data_only=True)
    data = {}

    for sheet in TAKENAKA_SHEETS:
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

def classify_action(src):
    if not src:
        return "NOT FOUND IN TAKENAKA SOURCE"

    s1 = norm_upper(src.get("Takenaka Status 1"))
    s2 = norm_upper(src.get("Takenaka Status 2"))
    s3 = norm_upper(src.get("Takenaka Status 3"))

    if s1 in ["CLOSED", "CLOSE"]:
        return "UPDATE TRACKING TO CLOSED"
    if "RETURNED" in s1 or "RETURNED" in s2:
        return "RETURNED BY NV5 / NEED RESUBMIT"
    if "OVERDUE" in s3:
        return "OVERDUE / FOLLOW UP"
    if "OPEN" in s1 and ("ON PROCESS" in s2 or "ON PROGRESS" in s2):
        return "OPEN & ON PROCESS"
    if "OPEN" in s1:
        return "OPEN"
    return "CHECK"

def should_include_tracking(status, info):
    status_u = norm_upper(status)
    info_u = norm_upper(info)

    return (
        "OPEN" in status_u
        or "ON PROGRESS" in status_u
        or "ON PROCESS" in status_u
        or "ON PROGRESS" in info_u
        or "ON PROCESS" in info_u
    )

def generate_report(tracking_file, takenaka_file):
    takenaka_map = read_takenaka(takenaka_file)

    read_wb = load_workbook(tracking_file, data_only=True)
    output_wb = load_workbook(tracking_file)

    report_sheet = "Open_On_Process_Compare"
    if report_sheet in output_wb.sheetnames:
        del output_wb[report_sheet]

    report_ws = output_wb.create_sheet(report_sheet)

    headers_out = [
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
    report_ws.append(headers_out)

    fills = {
        "UPDATE TRACKING TO CLOSED": PatternFill(fill_type="solid", fgColor="C6EFCE"),
        "OPEN & ON PROCESS": PatternFill(fill_type="solid", fgColor="FFEB9C"),
        "OPEN": PatternFill(fill_type="solid", fgColor="BDD7EE"),
        "OVERDUE / FOLLOW UP": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "RETURNED BY NV5 / NEED RESUBMIT": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "NOT FOUND IN TAKENAKA SOURCE": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "CHECK": PatternFill(fill_type="solid", fgColor="D9EAF7"),
    }

    for cell in report_ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    rows = []
    total_docs = 0
    focus_docs = 0

    for sheet in TRACKING_SHEETS:
        if sheet not in read_wb.sheetnames:
            continue

        ws = read_wb[sheet]

        header_row, headers = find_header_row_and_columns(ws)
        doc_no_col = get_col(headers, ["Document No.", "Document No", "Ref No.", "Ref No", "Drawing ref No."])
        doc_name_col = get_col(headers, ["Document Name", "Equipment Name", "Description"])
        status_col = get_col(headers, ["Status"])
        info_col = get_col(headers, ["Info"])

        if not doc_no_col or not status_col or not info_col:
            continue

        for row in range(header_row + 1, ws.max_row + 1):
            doc_no = ws.cell(row=row, column=doc_no_col).value
            doc_name = ws.cell(row=row, column=doc_name_col).value if doc_name_col else ""
            tracking_status = ws.cell(row=row, column=status_col).value
            info = ws.cell(row=row, column=info_col).value

            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            total_docs += 1

            if not should_include_tracking(tracking_status, info):
                continue

            focus_docs += 1
            key = base_doc_no(doc_no)
            src = takenaka_map.get(key)
            action = classify_action(src)
            checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

            if src:
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
            report_ws[f"K{report_ws.max_row}"].fill = fills.get(action, fills["CHECK"])

            rows.append({
                "Tracking Sheet": new_row[0],
                "Document No": new_row[1],
                "Document Name": new_row[2],
                "Tracking Status": new_row[3],
                "Info": new_row[4],
                "Takenaka Sheet": new_row[5],
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
    output_wb.save(output)
    output.seek(0)

    return output, total_docs, focus_docs, rows

# =============================
# SESSION STATE
# =============================
defaults = {
    "result_df": None,
    "report": None,
    "total_docs": 0,
    "focus_docs": 0,
    "action_counts": {},
    "last_updated": None,
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# =============================
# HERO
# =============================
st.markdown("""
<div class="hero">
    <div class="hero-grid">
        <div>
            <div class="hero-title">💎 Topaz Smart Document Tracker</div>
            <div class="hero-sub">TOPAZ BKK1 | ICT Document Control Dashboard</div>
        </div>
        <div class="pill">Clean Dashboard</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================
# UPLOAD AREA
# =============================
if st.session_state.role == "admin":
    st.markdown('<div class="admin">👩‍💼 Admin mode: Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

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
        if not df.empty:
            df["Action"] = df["Action"].apply(normalize_action)

        action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}

        st.session_state.result_df = df
        st.session_state.report = report
        st.session_state.total_docs = total_docs
        st.session_state.focus_docs = focus_docs
        st.session_state.action_counts = action_counts
        st.session_state.last_updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
        st.rerun()

else:
    st.markdown('<div class="notice">ℹ️ Viewer mode: only Pavinee can upload files and generate dashboard.</div>', unsafe_allow_html=True)

# =============================
# DASHBOARD
# =============================
if st.session_state.result_df is not None:
    df = st.session_state.result_df.copy()
    action_counts = st.session_state.action_counts

    open_process = action_counts.get("OPEN & ON PROCESS", 0)
    update_closed = action_counts.get("UPDATE TRACKING TO CLOSED", 0)
    returned = action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)
    overdue = action_counts.get("OVERDUE / FOLLOW UP", 0)
    not_found = action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)
    check = action_counts.get("CHECK", 0)
    open_only = action_counts.get("OPEN", 0)

    st.success(f"Dashboard generated successfully ✅ | Last updated: {st.session_state.last_updated}")

    k1, k2, k3, k4, k5 = st.columns(5)
    cards = [
        ("📁", "Total Documents", st.session_state.total_docs, "All DETH-NSC documents", "#7c3aed", "#ede9fe"),
        ("🕒", "Open / On Progress", st.session_state.focus_docs, "Require review", "#2563eb", "#dbeafe"),
        ("▶", "Open & On Process", open_process, "Waiting review", "#16a34a", "#dcfce7"),
        ("🔄", "Need Update", update_closed, "Update tracking", "#f97316", "#ffedd5"),
        ("⚠", "Overdue", overdue, "Follow up", "#ef4444", "#fee2e2"),
    ]

    for col, (icon, title, value, sub, accent, soft) in zip([k1, k2, k3, k4, k5], cards):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--accent:{accent}; --soft:{soft};">
                <div class="kpi-icon">{icon}</div>
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
                <div class="kpi-sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)

    st.write("")

    chart_col, quick_col = st.columns([2, 1])

    with chart_col:
        st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)

        summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])
        if not summary_df.empty:
            chart = (
                alt.Chart(summary_df)
                .mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8)
                .encode(
                    x=alt.X("Action:N", sort="-y", axis=alt.Axis(labelAngle=-30)),
                    y=alt.Y("Count:Q"),
                    tooltip=["Action", "Count"]
                )
                .properties(height=300)
            )
            st.altair_chart(chart, use_container_width=True)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.markdown('</div>', unsafe_allow_html=True)

    with quick_col:
        st.markdown('<div class="panel"><div class="panel-title">⚡ Quick Action</div>', unsafe_allow_html=True)

        quick_items = [
            ("Returned by NV5", returned),
            ("Need update closed", update_closed),
            ("Overdue follow up", overdue),
            ("Not found in Takenaka", not_found),
            ("Check manually", check),
            ("Open only", open_only),
        ]

        for label, count in quick_items:
            st.markdown(
                f'<div class="quick-row"><span>{label}</span><span class="count-pill">{count}</span></div>',
                unsafe_allow_html=True
            )

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)
# Ensure dataframe has Action column
if df.empty:
    df = pd.DataFrame(columns=[
        "Tracking Sheet",
        "Document No",
        "Document Name",
        "Tracking Status",
        "Info",
        "Takenaka Sheet",
        "Takenaka Status 1",
        "Takenaka Status 2",
        "Takenaka Status 3",
        "Action",
        "Checked Time",
    ])

if "Action" not in df.columns:
    df["Action"] = ""    
    f1, f2, f3 = st.columns([1, 2, 0.8])
    with f1:
action_list = ["All"]

    if "Action" in df.columns:
        actions = (
            df["Action"]
            .fillna("")
            .astype(str)
            .unique()
            .tolist()
        )
    
        actions = sorted([a for a in actions if a != ""])
        action_list.extend(actions)
    
    selected_action = st.selectbox(
        "Filter by Action",
        action_list
    )
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

    st.dataframe(filtered_df, use_container_width=True, height=480)

    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("Please upload both Excel files and click Generate Dashboard.")
