import streamlit as st

def apply_style():
    st.markdown(
        """
<style>
.stApp {
    background: #f4f7fb;
    color: #0f172a;
}

.block-container {
    max-width: 1500px;
    padding-top: 1.2rem;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #061124 0%, #0f172a 70%, #111827 100%);
}

[data-testid="stSidebar"] * {
    color: white;
}

.sidebar-logo-title {
    font-size: 31px;
    font-weight: 900;
    margin-top: 8px;
    letter-spacing: .5px;
}

.sidebar-subtitle {
    color: #c7d2fe;
    font-size: 13px;
    margin-bottom: 20px;
}

.sidebar-card {
    padding: 16px;
    border-radius: 18px;
    background: rgba(255,255,255,.08);
    border: 1px solid rgba(255,255,255,.14);
    margin: 16px 0;
}

.nav-active {
    padding: 11px 13px;
    border-radius: 13px;
    margin: 8px 0;
    background: linear-gradient(90deg,#4f46e5,#7c3aed);
    font-weight: 800;
}

.nav-item {
    padding: 11px 13px;
    border-radius: 13px;
    margin: 8px 0;
    background: rgba(255,255,255,.06);
    font-weight: 700;
}

.hero {
    background: linear-gradient(90deg, #071226 0%, #0b1b4d 50%, #172554 100%);
    border-radius: 26px;
    padding: 30px 34px;
    color: white;
    box-shadow: 0 20px 45px rgba(15, 23, 42, .22);
    margin-bottom: 22px;
}

.hero-grid {
    display: grid;
    grid-template-columns: auto 1fr auto;
    gap: 22px;
    align-items: center;
}

.hero-logo {
    width: 92px;
    height: 92px;
    object-fit: contain;
}

.hero-title {
    font-size: 40px;
    font-weight: 950;
    line-height: 1.05;
    margin-bottom: 8px;
}

.hero-subtitle {
    color: #dbeafe;
    font-size: 15px;
}

.hero-badge {
    padding: 12px 18px;
    border-radius: 999px;
    background: rgba(255,255,255,.13);
    border: 1px solid rgba(255,255,255,.22);
    font-weight: 800;
    white-space: nowrap;
}

.info-box {
    padding: 16px 18px;
    border-radius: 18px;
    background: linear-gradient(90deg, #ecfdf5, #ffffff);
    border: 1px solid #bbf7d0;
    color: #166534;
    font-weight: 800;
    margin: 18px 0;
}

.viewer-box {
    padding: 16px 18px;
    border-radius: 18px;
    background: linear-gradient(90deg, #eff6ff, #ffffff);
    border: 1px solid #bfdbfe;
    color: #1d4ed8;
    font-weight: 800;
    margin: 18px 0;
}

.warning-box {
    padding: 16px 18px;
    border-radius: 18px;
    background: linear-gradient(90deg, #fff7ed, #ffffff);
    border: 1px solid #fed7aa;
    color: #9a3412;
    font-weight: 800;
    margin: 18px 0;
}

.kpi-card {
    padding: 24px;
    border-radius: 22px;
    background: white;
    box-shadow: 0 14px 32px rgba(15, 23, 42, 0.08);
    border: 1px solid #e5e7eb;
    min-height: 150px;
    position: relative;
    overflow: hidden;
}

.kpi-card:after {
    content:"";
    position:absolute;
    left:0;
    right:0;
    bottom:0;
    height:5px;
    background: var(--accent);
}

.kpi-icon {
    font-size: 26px;
    margin-bottom: 14px;
}

.kpi-title {
    font-size: 13px;
    color: #64748b;
    font-weight: 850;
}

.kpi-value {
    font-size: 38px;
    font-weight: 950;
    color: #0f172a;
    line-height: 1.1;
    margin: 7px 0;
}

.kpi-sub {
    color: #64748b;
    font-size: 13px;
}

.panel {
    background: white;
    border:1px solid #e2e8f0;
    border-radius: 24px;
    box-shadow: 0 14px 32px rgba(15,23,42,.07);
    padding: 24px;
    margin-bottom: 18px;
}

.panel-title {
    font-size: 22px;
    font-weight: 950;
    color:#0f172a;
    margin-bottom: 16px;
}

.quick-row {
    display:flex;
    justify-content:space-between;
    align-items:center;
    padding:13px 14px;
    border-radius:14px;
    background:#f8fafc;
    border:1px solid #e2e8f0;
    margin-bottom:10px;
}

.count-pill {
    padding:4px 11px;
    border-radius:999px;
    font-weight:900;
    background:#dbeafe;
    color:#1d4ed8;
}

[data-testid="stFileUploaderDropzone"] {
    border-radius: 18px;
    border: 1px solid #dbe3ef;
    background: #f8fafc;
}

.stButton > button {
    border-radius: 14px;
    font-weight: 800;
}
</style>
""",
        unsafe_allow_html=True,
    )
