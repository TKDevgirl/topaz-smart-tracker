# Topaz Smart Tracker V7 Enterprise

Enterprise-style Streamlit dashboard for Topaz BKK1 ICT Document Control.

## Main Features

- Admin / Viewer mode
- Shared dashboard storage
- Tracking document upload
- Takenaka comparison
- RFA / RFI latest revision summary
- RFA / RFI category summary
- Action summary
- Excel export
- Modular code structure

## Project Structure

```text
topaz_smart_tracker_v7_enterprise/
├── app.py
├── config/
│   ├── settings.py
│   └── columns.py
├── core/
│   ├── session.py
│   ├── storage.py
│   └── utils.py
├── services/
│   ├── tracking_service.py
│   ├── takenaka_service.py
│   ├── summary_service.py
│   └── report_service.py
├── ui/
│   ├── theme.py
│   ├── sidebar.py
│   ├── layout.py
│   ├── cards.py
│   └── tables.py
├── assets/
├── data/
└── tests/
```

## Run

```bash
streamlit run app.py
```

## Login

- Admin usernames: `admin`, `pavinee`
- Other usernames become Viewer
