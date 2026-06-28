import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from io import BytesIO
import pandas as pd
import base64
import os
import json
from pathlib import Path

st.set_page_config(page_title="Topaz Smart Document Tracker", page_icon="💎", layout="wide")

tracking_sheets = ["RFA", "RFI"]
takenaka_sheets = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

LOGO_PATH = "assets/logo.png"
DATA_DIR = Path("data")
LATEST_CSV = DATA_DIR / "latest_result.csv"
LATEST_META = DATA_DIR / "latest_meta.json"
LATEST_REPORT = DATA_DIR / "latest_report.xlsx"
DATA_DIR.mkdir(exist_ok=True)

# =========================
# STYLE - UI ONLY
# =========================
st.markdown("""
<style>
.stApp { background: #f4f7fb; color: #0f172a; }
.block-container { max-width: 1500px; padding-top: 1.2rem; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #061124 0%, #0f172a 70%, #111827 100%); }
[data-testid="stSidebar"] * { color: white; }
.sidebar-logo-title { font-size: 31px; font-weight: 900; margin-top: 8px; letter-spacing: .5px; }
.sidebar-subtitle { color: #c7d2fe; font-size: 13px; margin-bottom: 20px; }
.sidebar-card { padding: 16px; border-radius: 18px; background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.14); margin: 16px 0; }
.nav-active { padding: 11px 13px; border-radius: 13px; margin: 8px 0; background: linear-gradient(90deg,#4f46e5,#7c3aed); font-weight: 800; }
.nav-item { padding: 11px 13px; border-radius: 13px; margin: 8px 0; background: rgba(255,255,255,.06); font-weight: 700; }
.hero { background: linear-gradient(90deg, #071226 0%, #0b1b4d 50%, #172554 100%); border-radius: 26px; padding: 30px 34px; color: white; box-shadow: 0 20px 45px rgba(15, 23, 42, .22); margin-bottom: 22px; }
.hero-grid { display: grid; grid-template-columns: auto 1fr auto; gap: 22px; align-items: center; }
.hero-logo { width: 86px; height: 86px; object-fit: contain; }
.hero-title { font-size: 40px; font-weight: 950; line-height: 1.05; margin-bottom: 8px; }
.hero-subtitle { color: #dbeafe; font-size: 15px; }
.hero-badge { padding: 12px 18px; border-radius: 999px; background: rgba(255,255,255,.13); border: 1px solid rgba(255,255,255,.22); font-weight: 800; white-space: nowrap; }
.info-box { padding: 16px 18px; border-radius: 18px; background: linear-gradient(90deg, #ecfdf5, #ffffff); border: 1px solid #bbf7d0; color: #166534; font-weight: 800; margin: 18px 0; }
.viewer-box { padding: 16px 18px; border-radius: 18px; background: linear-gradient(90deg, #eff6ff, #ffffff); border: 1px solid #bfdbfe; color: #1d4ed8; font-weight: 800; margin: 18px 0; }
.warning-box { padding: 16px 18px; border-radius: 18px; background: linear-gradient(90deg, #fff7ed, #ffffff); border: 1px solid #fed7aa; color: #9a3412; font-weight: 800; margin: 18px 0; }
.kpi-card { padding: 24px; border-radius: 22px; background: white; box-shadow: 0 14px 32px rgba(15, 23, 42, 0.08); border: 1px solid #e5e7eb; min-height: 150px; position: relative; overflow: hidden; }
.kpi-card:after { content:""; position:absolute; left:0; right:0; bottom:0; height:5px; background: var(--accent); }
.kpi-icon { font-size: 26px; margin-bottom: 14px; }
.kpi-title { font-size: 13px; color: #64748b; font-weight: 850; }
.kpi-value { font-size: 38px; font-weight: 950; color: #0f172a; line-height: 1.1; margin: 7px 0; }
.kpi-sub { color: #64748b; font-size: 13px; }
.panel { background: white; border:1px solid #e2e8f0; border-radius: 24px; box-shadow: 0 14px 32px rgba(15,23,42,.07); padding: 24px; margin-bottom: 18px; }
.panel-title { font-size: 22px; font-weight: 950; color:#0f172a; margin-bottom: 16px; }
.quick-row { display:flex; justify-content:space-between; align-items:center; padding:13px 14px; border-radius:14px; background:#f8fafc; border:1px solid #e2e8f0; margin-bottom:10px; }
.count-pill { padding:4px 11px; border-radius:999px; font-weight:900; background:#dbeafe; color:#1d4ed8; }
[data-testid="stFileUploaderDropzone"] { border-radius: 18px; border: 1px solid #dbe3ef; background: #f8fafc; }
.stButton > button { border-radius: 14px; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

def image_to_base64(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def render_logo_html(size_class="hero-logo"):
    logo64 = image_to_base64(LOGO_PATH)
    if logo64:
        return f'<img class="{size_class}" src="data:image/png;base64,{logo64}">'
    return '<div style="font-size:62px;">💎</div>'

def save_latest_dashboard(df, report, total_docs, open_docs, action_counts):
    DATA_DIR.mkdir(exist_ok=True)
    df.to_csv(LATEST_CSV, index=False, encoding="utf-8-sig")
    with open(LATEST_REPORT, "wb") as f:
        f.write(report.getvalue())
    meta = {
        "total_docs": int(total_docs),
        "open_docs": int(open_docs),
        "action_counts": action_counts,
        "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
    }
    with open(LATEST_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def load_latest_dashboard():
    if not LATEST_CSV.exists() or not LATEST_META.exists():
        return None, None, None
    df = pd.read_csv(LATEST_CSV)
    with open(LATEST_META, "r", encoding="utf-8") as f:
        meta = json.load(f)
    report_bytes = LATEST_REPORT.read_bytes() if LATEST_REPORT.exists() else None
    return df, meta, report_bytes

def hydrate_session_from_latest():
    if st.session_state.result_df is not None:
        return
    df, meta, report_bytes = load_latest_dashboard()
    if df is None or meta is None:
        return
    st.session_state.result_df = df
    st.session_state.report = report_bytes
    st.session_state.total_docs = int(meta.get("total_docs", 0))
    st.session_state.open_docs = int(meta.get("open_docs", 0))
    st.session_state.action_counts = meta.get("action_counts", {})
    st.session_state.last_updated = meta.get("last_updated", "")

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
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=145)
    else:
        st.markdown("## 💎")
    st.markdown('<div class="sidebar-logo-title">TOPAZ</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-subtitle">Smart Document Tracker</div>', unsafe_allow_html=True)
    st.divider()

    if not st.session_state.logged_in:
        username = st.text_input("Username")
        if st.button("Login", use_container_width=True):
            st.session_state.username = username.strip()
            if username.strip().lower() in ["pavinee", "admin"]:
                st.session_state.role = "admin"
            else:
                st.session_state.role = "viewer"
            st.session_state.logged_in = True
            st.rerun()
    else:
        st.markdown(f"""<div class="sidebar-card"><b>🔑 Role</b><br>{st.session_state.role.title()}</div>""", unsafe_allow_html=True)
        if st.button("Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.role = "viewer"
            st.session_state.username = ""
            st.rerun()

    st.divider()
    st.markdown('<div class="nav-active">🏠 Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-item">⬇ Download Report</div>', unsafe_allow_html=True)
    st.divider()
    st.caption("Topaz Smart Document Tracker V6 Shared")

# =========================
# ORIGINAL LOGIC KEPT
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
            key = base_doc_no(doc_no)
            data[key] = {
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
    headers = ["Tracking Sheet","Document No","Document Name","Tracking Status","Info","Takenaka Sheet","Takenaka Doc No","Takenaka Status 1","Takenaka Status 2","Takenaka Status 3","Action","Checked Time"]
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
            tracking_status = ws[f"F{row}"].value
            info = ws[f"G{row}"].value
            if not doc_no:
                continue
            total_docs += 1
            tracking_status_text = str(tracking_status or "").strip().upper()
            info_text = str(info or "").strip().upper()
            if ("OPEN" not in tracking_status_text and "ON PROGRESS" not in tracking_status_text and "ON PROCESS" not in tracking_status_text and "ON PROGRESS" not in info_text and "ON PROCESS" not in info_text):
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
                    action, fill = "UPDATE TRACKING TO CLOSED", green
                elif s1 == "RETURNED":
                    action, fill = "RETURNED BY NV5 / NEED RESUBMIT", red
                elif s3 == "OVERDUE":
                    action, fill = "OVERDUE / FOLLOW UP", red
                elif s1 == "OPEN" and s2 == "ON PROCESS":
                    action, fill = "OPEN & ON PROCESS", yellow
                elif s1 == "OPEN":
                    action, fill = "OPEN", blue
                else:
                    action, fill = "CHECK", blue
                new_row = [sheet, doc_no, doc_name, tracking_status, info, src["Takenaka Sheet"], src["Takenaka Doc No"], src["Takenaka Status 1"], src["Takenaka Status 2"], src["Takenaka Status 3"], action, checked_time]
            else:
                action, fill = "NOT FOUND IN TAKENAKA SOURCE", red
                new_row = [sheet, doc_no, doc_name, tracking_status, info, "", "", "", "", "", action, checked_time]

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

# =========================
# SESSION STATE
# =========================
if "result_df" not in st.session_state:
    st.session_state.result_df = None
    st.session_state.report = None
    st.session_state.total_docs = 0
    st.session_state.open_docs = 0
    st.session_state.action_counts = {}
    st.session_state.last_updated = ""

hydrate_session_from_latest()

# =========================
# HEADER
# =========================
st.markdown(f"""
<div class="hero">
    <div class="hero-grid">
        <div>{render_logo_html()}</div>
        <div>
            <div class="hero-title">Topaz Smart Document Tracker</div>
            <div class="hero-subtitle">TOPAZ BKK1 | ICT Document Control Dashboard</div>
        </div>
        <div class="hero-badge">Shared Dashboard</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================
# UPLOAD AREA
# =========================
if st.session_state.role == "admin":
    st.markdown('<div class="info-box">👩‍💼 Admin mode: Admin can upload files, generate dashboard, and update shared data.</div>', unsafe_allow_html=True)
    col_upload_1, col_upload_2, col_button = st.columns([1, 1, 0.8])
    with col_upload_1:
        tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])
    with col_upload_2:
        takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])
    with col_button:
        st.write("")
        st.write("")
        generate_clicked = st.button("🚀 Generate Dashboard", type="primary", use_container_width=True)

    if tracking_file and takenaka_file and generate_clicked:
        with st.spinner("Reading files and generating dashboard..."):
            report, total_docs, open_docs, rows = generate_report(tracking_file, takenaka_file)

        df = pd.DataFrame(rows)
        action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}
        save_latest_dashboard(df, report, total_docs, open_docs, action_counts)

        st.session_state.result_df = df
        st.session_state.report = report
        st.session_state.total_docs = total_docs
        st.session_state.open_docs = open_docs
        st.session_state.action_counts = action_counts
        st.session_state.last_updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
        st.rerun()
else:
    st.markdown('<div class="viewer-box">ℹ️ Viewer mode: dashboard is read-only. Only Admin can upload files and update shared data.</div>', unsafe_allow_html=True)

# =========================
# DASHBOARD
# =========================
if st.session_state.result_df is not None:
    df = st.session_state.result_df
    action_counts = st.session_state.action_counts
    last_updated = st.session_state.get("last_updated", "")

    st.success(f"Dashboard loaded successfully ✅ | Last updated: {last_updated}" if last_updated else "Dashboard loaded successfully ✅")

    c1, c2, c3, c4, c5 = st.columns(5)
    cards = [
        ("📁", "Total Documents", st.session_state.total_docs, "All documents", "#7c3aed"),
        ("🕒", "Open / On Progress", st.session_state.open_docs, "Documents", "#2563eb"),
        ("▶️", "Open & On Process", action_counts.get("OPEN & ON PROCESS", 0), "Waiting review", "#16a34a"),
        ("🔄", "Need Update", action_counts.get("UPDATE TRACKING TO CLOSED", 0), "Update tracking", "#f97316"),
        ("⚠️", "Overdue", action_counts.get("OVERDUE / FOLLOW UP", 0), "Follow up", "#ef4444"),
    ]
    for col, (icon, title, value, sub, color) in zip([c1,c2,c3,c4,c5], cards):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--accent:{color};">
                <div class="kpi-icon">{icon}</div>
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
                <div class="kpi-sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)

    st.write("")
    st.write("")

    chart_col, table_col = st.columns([1.4, 1])
    with chart_col:
        st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)
        summary_df = pd.DataFrame([{"Action": k, "Count": v} for k, v in action_counts.items()])
        if not summary_df.empty:
            st.bar_chart(summary_df.set_index("Action"))
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        else:
            st.info("No action data found.")
        st.markdown("</div>", unsafe_allow_html=True)

    with table_col:
        st.markdown('<div class="panel"><div class="panel-title">🧭 Quick Action</div>', unsafe_allow_html=True)
        quick_items = [
            ("Returned by NV5", action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)),
            ("Need update closed", action_counts.get("UPDATE TRACKING TO CLOSED", 0)),
            ("Overdue follow up", action_counts.get("OVERDUE / FOLLOW UP", 0)),
            ("Not found in Takenaka", action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)),
        ]
        for label, count in quick_items:
            st.markdown(f'<div class="quick-row"><span>{label}</span><span class="count-pill">{count}</span></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)
    filter_col, search_col, download_col = st.columns([1, 2, 0.8])
    with filter_col:
        selected_action = st.selectbox("Filter by Action", ["All"] + sorted(df["Action"].dropna().unique().tolist()))
    with search_col:
        search = st.text_input("Search Document No / Document Name")

    filtered_df = df.copy()
    if selected_action != "All":
        filtered_df = filtered_df[filtered_df["Action"] == selected_action]
    if search:
        filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)]

    st.dataframe(filtered_df, use_container_width=True, height=480)

    with download_col:
        st.write("")
        st.write("")
        if st.session_state.report is not None:
            st.download_button(label="⬇️ Download Excel Report", data=st.session_state.report, file_name="Open_On_Process_Compare.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.markdown('<div class="warning-box">⚠️ No shared dashboard data available yet. Please login as Admin and generate dashboard once.</div>', unsafe_allow_html=True)
