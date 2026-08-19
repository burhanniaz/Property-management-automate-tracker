# property-tracker-backend

Small Express API that reads the Property Management Tracker Google Sheet's
live Dashboard KPIs (the `KPI_*`, `PM_Workload_Table`, and `Heartbeat_Cell`
named ranges created by `../build_dashboard.py`) and serves them as JSON.

It discovers named ranges dynamically — the same design as
`../AppsScript/Code.gs` — so adding or renaming a `KPI_*` named range in the
sheet shows up here automatically, nothing is hardcoded to a specific KPI.

## Setup

1. `npm install`
2. `cp .env.example .env` and fill in:
   - `SPREADSHEET_ID` — already defaulted to the current sheet's ID.
   - `GOOGLE_SERVICE_ACCOUNT_KEY_PATH` — path to your service account's
     JSON key (the one shared as Viewer on the sheet). Keep this file out
     of version control. On platforms with no file storage, use
     `GOOGLE_SERVICE_ACCOUNT_KEY_JSON` instead (the key's full JSON as one
     line) — see the comment in `.env.example`.
3. `npm start` (or `npm run dev` to auto-restart on file changes).

## Endpoints

- `GET /health` → `{ "status": "ok" }` — no Google auth required, useful
  for uptime checks.
- `GET /api/dashboard` → every tracked named range as one JSON object:
  ```json
  {
    "data": {
      "KPI_TotalUnits": "229",
      "KPI_OccupancyRate": "74.9%",
      "PM_Workload_Table": [["Assigned PM", "Active Units", ...], ...],
      "Heartbeat_Cell": "2026-08-19T05:44:03.163-04:00"
    },
    "generatedAt": "2026-08-19T05:50:00.000Z",
    "cached": true,
    "cacheAgeSeconds": 7
  }
  ```
- `GET /api/dashboard/:name` → one named range, e.g.
  `GET /api/dashboard/KPI_OccupancyRate`.

## Freshness

Responses are cached in memory for `CACHE_TTL_SECONDS` (default 20s) to
avoid re-hitting the Sheets API on every request. This doesn't limit
freshness in practice: the sheet itself only recalculates its `TODAY()`-based
KPIs (Leasing Pipeline, Growth) once a minute, driven by the
`../n8n/property-tracker-heartbeat.json` workflow — so anything faster than
that ceiling doesn't return newer data anyway.

## Auth model

Read-only via a Google service account (no user OAuth flow, no browser
consent needed). The service account must be shared as **Viewer** on the
Google Sheet. See the repo's top-level setup notes for how to create one.
