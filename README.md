# Topaz Smart Tracker V11.1 Polish Product

Enterprise Document Control Platform for ICT / Data Center projects.

## Purpose

Topaz Smart Tracker converts project document registers into an executive dashboard with:

- Document action tracking
- Latest revision logic
- RFA / RFI category summary
- Takenaka comparison
- SQLite database history
- Search and timeline
- Alert center
- Executive insights
- Export center

## What is improved in V11.1

- Cleaner dashboard flow
- Trend chart hidden until enough history exists
- More professional empty states
- Smaller Generate button
- Executive sections ordered for PM daily use
- Portfolio-ready README and release notes

## Architecture

```text
app.py
config/
core/
repositories/
services/
ui/
data/
docs/
```

## Run

```bash
streamlit run app.py
```

## Login

- Admin usernames: `admin`, `pavinee`
- Other usernames become Viewer

## Data Flow

```text
Excel Upload
    ↓
Generate Dashboard
    ↓
SQLite Database
    ↓
Executive Dashboard / Search / Timeline / Export
```

## Notes

V11.1 focuses on polish and stability rather than adding new heavy features.
