# DDO++ Attendance Office Agent

One Windows office PC agent (Python) that:

1. Optionally pulls attendance from **eSSL eTimeTrackLite** into Excel  
2. Reads that Excel (or any file dropped in the folder)  
3. Converts it to JSON and POSTs it to the **DDO++ API**

No frontend, database, or Node.js in this repo.

---

## How it works

```
eSSL (optional)  -->  D:\Attendance\{Month} {Location}.xls  -->  importer  -->  DDO++ API
scripts\essl_export.py     e.g. Sep Agra.xls           scripts\importer_script.py
```

- Each office PC has its own `LOCATION_CODE` in `.env` (`AGRA` / `NOIDA` / `HYD`).
- eSSL export saves Excel using **current month + location**, e.g. `Sep Agra.xls`, `Sep Noida.xls`, `Sep Hyd.xls`.
- HR can also drop Excel into the same folder (`ATTENDANCE_DIR`, default `D:\Attendance`). No office subfolders needed.
- The importer tags every file with that PC’s location and uploads it.

---

## Folder layout

```
Importer_script/
  README.md                 <-- this guide + Python package list
  Setup.bat                 <-- double-click install package
  Uninstall.bat             <-- remove schedulers + Start Menu shortcuts
  install.ps1               <-- setup engine used by Setup.bat
  .env.example              <-- copy to .env (never commit .env)
  scripts/
    essl_export.py          <-- automate eSSL -> Excel
    importer_script.py      <-- Excel -> JSON -> API
    scheduler.ps1           <-- Task Scheduler wrapper for importer
    install_daily_scheduler.ps1
    install_essl_export_scheduler.ps1
  processed/ failed/ logs/ output/   <-- runtime (local only)
```

### Python packages

`Setup.bat` installs these from this README (no separate `requirements.txt`):

```pip-requirements
xlrd>=2.0.1
openpyxl>=3.1.0
pywinauto>=0.6.8
comtypes>=1.4.0
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

### 2. One-click install package

```powershell
git clone <repo-url>
cd Importer_script
```

Then **double-click `Setup.bat`**.

That install package will:

1. Find Python, or install Python 3.12 with winget if missing  
2. Create runtime folders  
3. Create `.env` from `.env.example` if missing  
4. `pip install` the packages listed under **Python packages** above  
5. Verify packages import  
6. Create the attendance drop folder (`D:\Attendance` by default)  
7. Open Notepad so you can set `LOCATION_CODE` and the API token  
8. **Register both Task Scheduler jobs** (Mon–Fri 13:00):
   - `DDO-Attendance-Importer`
   - `DDO-eSSL-Export`
9. Add **Start Menu** shortcuts under `DDO++ Attendance Agent`

Skip scheduler registration if you only want packages/config:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -SkipSchedulers
```

To remove schedulers and Start Menu shortcuts later, double-click **`Uninstall.bat`** (keeps this folder and `.env`).

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

Optional eSSL settings (defaults match **this** office’s Device Management names):

```env
ESSL_USER=essl
ESSL_PASSWORD=essl
# Exact names from Device Management → Device Name column (comma-separated)
ESSL_DEVICES=LGF OUT,LGF IN,UGF OUT
# Always skip USB; also skip broken devices until fixed
ESSL_SKIP_DEVICES=USB,UGF IN 1
ESSL_REPORT=monthly-basic
# Company filter on Monthly Status Report (Deselect All, then this name)
ESSL_COMPANY=WalkingTree
```

**Per-branch devices:** each office edits `.env` only — do not hard-code other floors in the script.

| This office (example) | Meaning |
|-----------------------|---------|
| `LGF OUT` / `LGF IN` | Lower ground floor |
| `UGF OUT` / `UGF IN 1` | Upper ground floor (`UGF IN 1` skipped while buggy) |

Another branch with different readers just sets its own names, e.g.:

```env
ESSL_DEVICES=Floor1 IN,Floor1 OUT,Floor2 IN
ESSL_SKIP_DEVICES=USB
```

Names must match the **Device Name** column in eSSL exactly (spacing/spelling).

Notes:

- Export file name comes from `LOCATION_CODE` + current month (see below). You do not need `--export-name` for normal runs.
- When `UGF IN 1` is fixed here: remove it from `ESSL_SKIP_DEVICES` and add it to `ESSL_DEVICES`.

### 4. Create the drop folder

`Setup.bat` creates `ATTENDANCE_DIR` for you. To create it manually:

```powershell
mkdir D:\Attendance
```

Then put `.xls` / `.xlsx` files **directly** in that folder.

---

## Daily use

### Import Excel already in the drop folder

```powershell
python scripts\importer_script.py
```

One file:

```powershell
python scripts\importer_script.py --location NOIDA "C:\path\to\report.xls"
```

### Manual export from eSSL (recommended for smaller Excel)

On **Monthly Status Report** filter dialog, before Generate:

1. Report Type = **Basic Work Duration**
2. **From Date** = **To Date** = previous weekday (Mon–Fri only — Monday exports Friday, not Sunday; avoids weekend WO and overlapping “today’s In-only”)
3. Tick **Filter Company**
4. Click **Deselect All**
5. Click **WalkingTree** only
6. Click **Generate**
7. Export / Save As Excel into `D:\Attendance` (old file of the same name is moved to `D:\Attendance\previous\`)
   - Archived name looks like `Sep Agra_20260918_121824.xls`
   - The number is a timestamp: **`YYYYMMDD_HHMMSS`** when it was replaced  
     (e.g. `20260918_121824` = 18 Sep 2026 at 12:18:24). Not an employee/device ID.
8. Run importer: `python scripts\importer_script.py`

Filtering to Walking Tree keeps the file smaller and avoids API payload limits.

### Export from eSSL, then import (automated)

Default report is **Monthly Status → Basic Work Duration**, with **From=To=previous weekday** (skips Sat/Sun so Monday picks up Friday), **Filter Company → WalkingTree**, then Generate. Previous Excel of the same name is moved to `ATTENDANCE_DIR\previous\` before save (and Confirm Save As → Yes is clicked if Windows still asks).

Full flow (device sync + monthly basic export + save to `D:\Attendance` + logout + Close):

```powershell
python scripts\essl_export.py
python scripts\importer_script.py
```

Saves as **`{Month} {Location}.xls`** in `ATTENDANCE_DIR` (old same-name file goes to `previous\`):

| `LOCATION_CODE` | Example file (September) |
|-----------------|--------------------------|
| `AGRA` | `Sep Agra.xls` |
| `NOIDA` | `Sep Noida.xls` |
| `HYD` | `Sep Hyd.xls` |

When a new export would overwrite that name, the previous file is moved to:

`ATTENDANCE_DIR\previous\{Month} {Location}_YYYYMMDD_HHMMSS.xls`

Example: `D:\Attendance\previous\Sep Agra_20260918_121824.xls`  
→ archived on **2026-09-18** at **12:18:24** (local PC time). The suffix is only a replace timestamp so you can tell older copies apart.

`previous\` keeps at most **5** files (newest). Older archives are deleted automatically after each new export.

Useful variants:

```powershell
python scripts\essl_export.py --skip-sync          # export only, no device download
python scripts\essl_export.py --sync-only          # devices only
python scripts\essl_export.py --login-only         # test login
python scripts\essl_export.py --report monthly-basic   # default: importer-compatible
python scripts\essl_export.py --report detailed        # Daily Detailed (importer does not parse yet)
python scripts\essl_export.py --export-name "Custom Name.xls"   # override default name
```

**Devices synced by default:** `LGF OUT`, `LGF IN`, `UGF OUT`  
**Skipped:** `USB`, `UGF IN 1` (add `UGF IN 1` back in `.env` when the device is fixed)

**Important:** eSSL automation needs an **unlocked** interactive desktop (monitor off is OK; Win+L lock is not). Keep the eSSL window visible while it runs.

### Schedule (optional)

Importer (default Mon–Fri at 13:00 / 1 PM):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1
```

eSSL export (default Mon–Fri at 13:00 / 1 PM — session logged on + unlocked; monitor may be off):

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
2. Utilities → Device Management → select **LGF OUT / LGF IN / UGF OUT** (skip **USB** and **UGF IN 1**) → Start Download → wait  
3. Close dialogs until the main window is plain  
4. Attendance Reports → Monthly Reports → Monthly Status → **Report Type = Basic Work Duration**  
   then set **From Date = To Date = previous weekday** (Mon–Fri; Monday → Friday)
   then **Filter Company → Deselect All → WalkingTree → Generate**  
5. Export Excel into `ATTENDANCE_DIR` as `{Month} {Location}.xls` (previous file is moved to `ATTENDANCE_DIR\previous\`)  
6. Close dialogs → Log Off (3rd toolbar icon) → Close on login dialog  

App path used by default:

`C:\Program Files (x86)\essl\eTimeTrackLite\eTimeTrackLite.exe`

---

## Commit / share this project safely

**Do commit**

- `scripts/`, `README.md`, `.env.example`, `Setup.bat`, `Uninstall.bat`, `install.ps1`
- Folder placeholders (`.gitkeep` under `logs/`, `processed/`, etc.)
- Non-secret code

**Do not commit** (already ignored)

- `.env` (tokens and office location)
- Excel files (`*.xls`, `*.xlsx`)
- `logs/`, `output/` run data, `processed/`, `failed/`

`output/` keeps runtime state (`last_run.json`, `processed_hashes.json`, logs) plus the **newest 10** `batch_*` import JSON pairs. Older batches and leftover handoff/sample files are deleted automatically after each import. Override with `OUTPUT_BATCH_KEEP` in `.env` if needed.

`processed/` and `failed/` each keep only the **newest 15** Excel files (override with `ARCHIVE_KEEP`).

Typical commit:

```powershell
git status
git add README.md .env.example Setup.bat Uninstall.bat install.ps1 scripts .gitignore
git commit -m "Describe why you changed something, not only what files moved."
```

Never put real tokens in `.env.example` or in commit messages.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `xlrd` / `pywinauto` missing | Re-run `Setup.bat`, or `python -m pip install xlrd openpyxl pywinauto comtypes` with the same `python` you run scripts with |
| Python missing on a new PC | Run `Setup.bat` (uses winget), or install from https://www.python.org/downloads/ with PATH checked |
| Importer finds no files | Check `ATTENDANCE_DIR` exists and has `.xls`/`.xlsx`; check `LOCATION_CODE` |
| Wrong export name (e.g. `Sep Noida.xls` on Agra PC) | Set `LOCATION_CODE=AGRA` in `.env` and re-run; do not use a UTF-8 BOM-only broken `.env` |
| API / auth errors | Confirm token and `DDO_API_ENDPOINT` in `.env` |
| eSSL `Start Download` not found | Close other windows; leave eSSL visible; retry |
| Export file missing | Ensure `D:\Attendance` exists and is writable |
| `UGF IN 1` still selected | Confirm `ESSL_SKIP_DEVICES` includes `UGF IN 1` (default) |
| eSSL does nothing when PC locked | Unlock the session; schedule “Run only when user is logged on” |
| Wrong Python | Prefer `C:\Program Files\Python313\python.exe` over `py -3` free-threaded builds |

Debug UI dump (eSSL):

```powershell
python scripts\essl_export.py --dump-ui
```
