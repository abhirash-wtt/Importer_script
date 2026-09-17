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
| `LOCATION_CODE` | **This PC's office:** `AGRA`, `NOIDA`, or `HYD` |
| `DDO_API_ENDPOINT` | Import API URL (same for all offices unless told otherwise) |
| `ATTENDANCE_INTEGRATION_TOKEN` | Shared import token from the server team |
| `ATTENDANCE_DIR` | **Single drop folder** (default `D:\Attendance`) |

Example for Hyderabad:

```env
LOCATION_CODE=HYD
DDO_API_ENDPOINT=https://ddoplusnodeapi.walkingtree.tech/api/attendance/import
ATTENDANCE_INTEGRATION_TOKEN=your-token-here
ATTENDANCE_DIR=D:\Attendance
```

`DDO_API_TOKEN` is an optional fallback name for the same token.

## One drop folder (per PC)

1. Set `LOCATION_CODE` and `ATTENDANCE_DIR` in `.env`
2. Create the folder once: `mkdir D:\Attendance` (or whatever you set)
3. HR / anyone drops the Excel (`.xls` / `.xlsx`) **directly into that folder**
4. The scheduler converts it to JSON and POSTs it using this PC's `LOCATION_CODE`

No `NOIDA` / `AGRA` / `HYD` subfolders under `D:\Attendance` are required.

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

- Old backend/API/Postgres samples: `archive/unused/`
- Old TimeTrak RPA (AutoHotkey, export watcher): `archive/unused/timetrak_automation/`

Not required to run the importer.
