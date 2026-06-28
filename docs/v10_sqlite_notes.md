# V10 SQLite Notes

## Why SQLite?

SQLite keeps the latest dashboard and all historical runs in one database file:

```text
data/topaz_tracker.db
```

## Tables

- `dashboard_runs`
- `action_list`
- `status_summary`
- `status_detail`

## Important

Users do not manually edit SQL.

The data changes when Admin uploads the latest Excel files and clicks Generate Dashboard.
