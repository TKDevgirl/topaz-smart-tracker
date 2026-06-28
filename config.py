from pathlib import Path

APP_TITLE = "Topaz Smart Document Tracker"
APP_VERSION = "V6.1.8 Structured"

TRACKING_SHEETS = ["RFA", "RFI"]
TAKENAKA_SHEETS = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

LOGO_PATH = "assets/logo.png"

DATA_DIR = Path("data")
LATEST_CSV = DATA_DIR / "latest_result.csv"
LATEST_META = DATA_DIR / "latest_meta.json"
LATEST_REPORT = DATA_DIR / "latest_report.xlsx"
LATEST_STATUS_SUMMARY = DATA_DIR / "latest_status_summary.csv"
LATEST_STATUS_DETAIL = DATA_DIR / "latest_status_detail.csv"

DATA_DIR.mkdir(parents=True, exist_ok=True)
