from pathlib import Path

APP_TITLE = "Topaz Smart Tracker V10"
APP_VERSION = "V10 SQLite Edition"
PROJECT_NAME = "TOPAZ BKK1"
APP_ICON = "💎"

TRACKING_SHEETS = ["RFA", "RFI"]
TAKENAKA_SHEETS = ["MAT_ICT", "DWG_ICT", "MTS_ICT"]

ADMIN_USERS = ["admin", "pavinee"]

LOGO_PATH = "assets/logo.png"

DATA_DIR = Path("data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

LATEST_RESULT_CSV = DATA_DIR / "latest_result.csv"
LATEST_META_JSON = DATA_DIR / "latest_meta.json"
LATEST_REPORT_XLSX = DATA_DIR / "latest_report.xlsx"
LATEST_STATUS_SUMMARY_CSV = DATA_DIR / "latest_status_summary.csv"
LATEST_STATUS_DETAIL_CSV = DATA_DIR / "latest_status_detail.csv"

HISTORY_DIR = DATA_DIR / "history"
HISTORY_DIR.mkdir(parents=True, exist_ok=True)

DATABASE_PATH = DATA_DIR / "topaz_tracker.db"
