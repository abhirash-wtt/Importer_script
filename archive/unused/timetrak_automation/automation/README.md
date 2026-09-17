# eSSL eTimeTrackLite automation steps

> **Status:** RPA scheduled task `DDO-TimeTrak-AutoExport` is **disabled** until device checkbox selection is reliable.
> Manual Excel export + `watch_exports.ps1` / daily importer still work. Re-enable with:
> `.\install_timetrack_rpa.ps1 -Enable` after `--select-only` shows LGF/UGF checked and USB unchecked.

## Exact office flow (confirmed)

1. **Login** — `essl` / `essl`
2. **Utilities → Device Management**
   - Leave **USB** unchecked
   - Check **LGF OUT, LGF IN, UGF OUT, UGF IN 1**
   - Click **Start Download**
   - Wait until sync finishes (Status / Logs Downloaded)
3. **Attendance Reports → Daily Attendance Reports → Basic Report**
   - Filter window opens
   - Set report type to **Basic Report** (not Summary)
   - Click **Generate**
4. Report viewer → toolbar **floppy/export** → **Excel**
5. **Save As** into `D:\Attendance\exports`

Then `watch_exports.ps1` + importer move/convert/POST the file.

## How to test in pieces

```powershell
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
$script = "D:\Importer_script\automation\TimeTrakExport.ahk"

# 1) Login only
& $ahk $script --login-only

# 2) Login + device sync only
& $ahk $script --sync-only

# 3) Full run (shorter sync wait)
& $ahk $script --test
```

Watch the screen. If a step misses, tell which step failed and we adjust clicks.
