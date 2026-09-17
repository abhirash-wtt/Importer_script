# Archived TimeTrak automation (not used)

Moved out of the active importer package. The live flow is:

1. Drop Excel into `ATTENDANCE_DIR` (e.g. `D:\Attendance`)
2. `scheduler.ps1` / `DDO-Attendance-Importer` converts and POSTs

## Contents

| Path | What it was |
|---|---|
| `automation/TimeTrakExport.ahk` | AutoHotkey RPA for eTimeTrackLite |
| `automation/capture_screen.ps1` | Screenshot helper for AHK |
| `automation/inspect_devices.py` | UI inspect helper |
| `install_timetrack_rpa.ps1` | Scheduled `DDO-TimeTrak-AutoExport` |
| `install_export_watcher.ps1` | Scheduled export folder watcher |
| `watch_exports.ps1` | Move exports into drop folder |

The `DDO-TimeTrak-AutoExport` task has been removed from Task Scheduler.
Do not re-register these unless you intentionally revive RPA work.
