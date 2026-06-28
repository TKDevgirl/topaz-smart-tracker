# Topaz Smart Tracker V10 SQLite Edition

Enterprise-style Streamlit dashboard for Topaz BKK1 ICT Document Control.

## V10 Highlights

- SQLite database storage
- `data/topaz_tracker.db`
- Every Generate creates a new `run_id`
- Dashboard loads the latest run by default
- Historical snapshots can be loaded from SQLite
- Excel upload still controls the dashboard data
- If Excel changes, upload and Generate again to update SQLite

## Main Flow

```text
Tracking_document.xlsx
Takenaka Summary.xlsx
        ↓
Generate Dashboard
        ↓
SQLite Database
        ↓
Dashboard / Search / Timeline / Analytics
```

## Run

```bash
streamlit run app.py
```

## Login

- Admin usernames: `admin`, `pavinee`
- Other usernames become Viewer
