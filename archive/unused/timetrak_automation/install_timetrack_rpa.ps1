# Schedules AutoHotkey TimeTrak export, then the normal importer can pick up the file.
#
# DISABLED BY DEFAULT until device checkbox selection is fully working.
# Use -Enable only after automation\TimeTrakExport.ahk is calibrated and verified.
#
# Prerequisites:
#   - AutoHotkey v2 installed
#   - automation\TimeTrakExport.ahk calibrated on this PC
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_timetrack_rpa.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_timetrack_rpa.ps1 -Enable
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_timetrack_rpa.ps1 -Unregister

param(
    [string]$TaskName = "DDO-TimeTrak-AutoExport",
    [string[]]$DailyTimes = @("09:00", "17:00"),
    [switch]$Enable,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script = Join-Path $Root "automation\TimeTrakExport.ahk"
if (-not (Test-Path $Script)) {
    throw "Missing $Script"
}

$ahk = @(
    "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe",
    "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey32.exe",
    "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey64.exe",
    "$env:ProgramFiles\AutoHotkey\v2\AutoHotkey32.exe",
    "$env:ProgramFiles\AutoHotkey\AutoHotkey.exe",
    "${env:ProgramFiles(x86)}\AutoHotkey\AutoHotkey.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $ahk) {
    throw "AutoHotkey v2 was not found. Install it from https://www.autohotkey.com/ then re-run."
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    exit 0
}

$action = New-ScheduledTaskAction -Execute $ahk -Argument "`"$Script`"" -WorkingDirectory (Split-Path $Script)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew
$triggers = foreach ($time in $DailyTimes) {
    New-ScheduledTaskTrigger -Daily -At $time
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings `
    -Description "RPA: export attendance Excel from eSSL TimeTrak Lite into D:\Attendance\exports (disabled until AHK device select works)" -Force | Out-Null

if ($Enable) {
    Enable-ScheduledTask -TaskName $TaskName | Out-Null
    Write-Host "Registered and ENABLED '$TaskName' at $($DailyTimes -join ', ')."
} else {
    Disable-ScheduledTask -TaskName $TaskName | Out-Null
    Write-Host "Registered '$TaskName' at $($DailyTimes -join ', ') but LEFT DISABLED."
    Write-Host "Re-run with -Enable only after TimeTrakExport.ahk checkbox selection is verified."
}
Write-Host "Also keep install_export_watcher.ps1 and install_daily_scheduler.ps1 enabled for manual Excel drops."
