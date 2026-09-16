# DDO++ Attendance Importer (office agent)

Office PC agent only: read attendance Excel (`.xls` / `.xlsx`), convert to JSON, and POST to the DDO++ API.

There is no frontend or database in this package.

## Requirements

- Windows office PC
- Python 3.10+
- Optional: Node.js 18+ (only if you use `attendance_import.js`)

## Quick install (new office / new branch checkout)

```powershell
git clone <repo-url>
cd Importer_script

# 1) Create local config (do not commit .env)
copy .env.example .env

# 2) Edit .env — set THIS office location and token
notepad .env

# 3) Python deps (main importer)
python -m pip install -r requirements.txt

# 4) Optional Node deps (JS importer only)
npm install
```

## Configure `.env` (required per office)

| Variable | What to set |
|---|---|
| `LOCATION_CODE` | **Must match this office:** `AGRA`, `NOIDA`, or `HYD` |
| `DDO_API_ENDPOINT` | Import API URL (same for all offices unless told otherwise) |
| `ATTENDANCE_INTEGRATION_TOKEN` | Shared import token from the server team |
| `ATTENDANCE_DIR` | Optional extra drop folder (default `D:\Attendance`) |

Example for Hyderabad:

```env
LOCATION_CODE=HYD
DDO_API_ENDPOINT=https://ddoplusnodeapi.walkingtree.tech/api/attendance/import
ATTENDANCE_INTEGRATION_TOKEN=your-token-here
ATTENDANCE_DIR=D:\Attendance
```

`DDO_API_TOKEN` is an optional fallback name for the same token.

## Drop folders

Put Excel reports in any of:

- `inbox\AGRA`, `inbox\NOIDA`, `inbox\HYD`
- `D:\Attendance` (or `D:\Attendance\AGRA`, `\NOIDA`, `\HYD`)

Files with the location in the name (for example `June Agra.xls`) also work at the root of those folders.

## Run once

```powershell
python importer_script.py --inbox
```

Or for a single file:

```powershell
python importer_script.py --location NOIDA "C:\path\to\report.xls"
```

## Schedule (optional)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install_daily_scheduler.ps1
```

## What stays local (not shared in git)

- `.env` (tokens / office location)
- Excel files (`*.xls`, `*.xlsx`)
- `logs/`, `output/` run state, `processed/`, `failed/`

## Archive

Old backend/API/Postgres/docs samples live under `archive/` and are not required to run the importer.
