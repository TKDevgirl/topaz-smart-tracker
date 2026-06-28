# Topaz Smart Tracker V8.0 AI Platform Starter

Enterprise-style Streamlit dashboard for Topaz BKK1 ICT Document Control.

## V8.0 Features

- Modular architecture from V7
- Admin / Viewer mode
- Shared dashboard storage
- Tracking document upload
- Takenaka comparison
- Latest revision logic
- RFA / RFI category summary
- Executive AI-style overview
- Project health score
- Recommended actions
- Category analytics
- History snapshots
- Excel export

## Project Structure

```text
topaz_smart_tracker_v8_ai_platform/
├── app.py
├── config/
├── core/
├── services/
│   ├── analytics_service.py
│   ├── history_service.py
│   ├── insight_service.py
│   ├── report_service.py
│   ├── summary_service.py
│   ├── takenaka_service.py
│   └── tracking_service.py
├── ui/
│   ├── executive.py
│   ├── history.py
│   ├── tables.py
│   └── ...
├── assets/
├── data/
├── docs/
└── tests/
```

## Run

```bash
streamlit run app.py
```

## Login

- Admin usernames: `admin`, `pavinee`
- Other usernames become Viewer

## Notes

This V8 starter uses AI-style deterministic insights generated from dashboard metrics.
It does not call an external AI API yet, so it is safe and stable for Streamlit deployment.
