# DDO++ Attendance Office Agent

One Windows office PC agent that:

1. Optionally pulls attendance from **eSSL eTimeTrackLite** into Excel  
2. Reads that Excel (or any file dropped in the folder)  
3. Converts it to JSON and POSTs it to the **DDO++ API**

No frontend and no database live in this repo.

---

## How it works

```
eSSL (optional)  -->  D:\Attendance\*.xls  -->  importer  -->  DDO++ API
scripts\essl_export.py     drop folder      scripts\importer_script.py
```

- Each office PC has its own `LOCATION_CODE` in `.env` (`AGRA` / `NOIDA` / `HYD`).
- HR drops Excel into one folder (`ATTENDANCE_DIR`, default `D:\Attendance`). No office subfolders needed.
- The importer tags every file with that PC’s location and uploads it.

---

## Folder layout

```
Importer_script/
  README.md                 <-- this guide (only docs you need)
  .env.example              <-- copy to .env (never commit .env)
  requirements.txt          <-- Python packages
  package.json              <-- optional Node importer
  scripts/
    essl_export.py          <-- automate eSSL -> Excel
    importer_script.py      <-- Excel -> JSON -> API (main)
    attendance_import.js    <-- optional JS version of importer
    scheduler.ps1           <-- Task Scheduler wrapper for importer
    install_daily_scheduler.ps1
    install_essl_export_scheduler.ps1
  inbox/ processed/ failed/ logs/ output/   <-- runtime (local only)
  archive/                  <-- old samples; not required to run
```

---

## Setup (new PC / new clone)

### 1. Prerequisites

| Need | Notes |
|------|--------|
| Windows PC | Office machine where Excel is dropped |
| Python 3.10+ | Use normal `python.exe` (not free-threaded `py -3` / `python3.13t`) |
| Git | To clone and commit |
| eTimeTrackLite | Only if you use auto-export from eSSL |
| Node.js 18+ | Optional; only for `attendance_import.js` |

### 2. Clone and install

```powershell
git clone <repo-url>
cd Importer_script

copy .env.example .env
notepad .env

python -m pip install -r requirements.txt

# Optional (JS importer only)
npm install
```

### 3. Edit `.env` (required)

| Variable | What to put |
|----------|-------------|
| `LOCATION_CODE` | This PC’s office: `AGRA`, `NOIDA`, or `HYD` |
| `DDO_API_ENDPOINT` | Import API URL from the server team |
| `ATTENDANCE_INTEGRATION_TOKEN` | Shared import token from the server team |
| `ATTENDANCE_DIR` | Drop folder, usually `D:\Attendance` |

Example:

```env
LOCATION_CODE=HYD
DDO_API_ENDPOINT=https://ddoplusnodeapi.walkingtree.tech/api/attendance/import
ATTENDANCE_INTEGRATION_TOKEN=your-token-here
ATTENDANCE_DIR=D:\Attendance
```

Optional eSSL settings (defaults usually work):

```env
ESSL_USER=essl
ESSL_PASSWORD=essl
ESSL_DEVICES=LGF OUT,LGF IN,UGF OUT,UGF IN 1
ESSL_SKIP_DEVICES=USB
ESSL_REPORT=detailed
```

### 4. Create the drop folder

```powershell
mkdir D:\Attendance
```

Then put `.xls` / `.xlsx` files **directly** in that folder.

---

## Daily use

### Import Excel already in the drop folder

```powershell
python scripts\importer_script.py --inbox
```

One file:

```powershell
python scripts\importer_script.py --location NOIDA "C:\path\to\report.xls"
```

### Export from eSSL, then import

Full flow (device sync + Daily Detailed report + save to `D:\Attendance` + logout + Close):

```powershell
python scripts\essl_export.py
python scripts\importer_script.py --inbox
```

Useful variants:

```powershell
python scripts\essl_export.py --skip-sync          # export only, no device download
python scripts\essl_export.py --sync-only          # devices only
python scripts\essl_export.py --login-only         # test login
python scripts\essl_export.py --report monthly-basic
```

**Important:** eSSL automation needs an **unlocked** interactive desktop (monitor off is OK; Win+L lock is not). Keep the eSSL window visible while it runs.

### Schedule (optional)

Importer (default 01:00 and 13:00):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1
```

eSSL export (default 07:30 — PC must stay unlocked):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1
```

Unregister:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1 -Unregister
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -Unregister
```

After a scheduled run, check:

- `output\last_run.txt`
- `output\scheduler_runs.log`
- `logs\essl_export.log` (for eSSL)

---

## What eSSL export does (short)

1. Login to eTimeTrackLite  
2. Utilities → Device Management → select LGF/UGF (skip USB) → Start Download → wait  
3. Close dialogs until the main window is plain  
4. Attendance Reports → Daily Attendance Reports → Detailed Attendance Report → Generate  
5. Export Excel into `ATTENDANCE_DIR`  
6. Close dialogs → Log Off (3rd toolbar icon) → Close on login dialog  

App path used by default:

`C:\Program Files (x86)\essl\eTimeTrackLite\eTimeTrackLite.exe`

---

## Commit / share this project safely

**Do commit**

- `scripts/`, `README.md`, `.env.example`, `requirements.txt`, `package.json`, `package-lock.json`
- Folder placeholders (`.gitkeep` under `inbox/`, `logs/`, etc.)
- Docs and non-secret code

**Do not commit** (already ignored)

- `.env` (tokens and office location)
- Excel files (`*.xls`, `*.xlsx`)
- `logs/`, `output/` run data, `processed/`, `failed/`
- `node_modules/`

Typical first commit from a clean clone after your changes:

```powershell
git status
git add README.md .env.example requirements.txt package.json package-lock.json scripts .gitignore
git commit -m "Describe why you changed something, not only what files moved."
```

Never put real tokens in `.env.example` or in commit messages.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `xlrd` / `pywinauto` missing | `python -m pip install -r requirements.txt` using the same `python` you run scripts with |
| Importer finds no files | Check `ATTENDANCE_DIR` exists and has `.xls`/`.xlsx`; check `LOCATION_CODE` |
| API / auth errors | Confirm token and `DDO_API_ENDPOINT` in `.env` |
| eSSL `Start Download` not found | Close other windows; leave eSSL visible; retry |
| Export file missing | Ensure `D:\Attendance` exists and is writable |
| eSSL does nothing when PC locked | Unlock the session; schedule “Run only when user is logged on” |
| Wrong Python | Prefer `C:\Program Files\Python313\python.exe` over `py -3` free-threaded builds |

Debug UI dump (eSSL):

```powershell
python scripts\essl_export.py --dump-ui
```

---

## Archive

`archive/samples/` and `archive/unused/` are old samples and experiments. You do not need them to set up or run the agent.
