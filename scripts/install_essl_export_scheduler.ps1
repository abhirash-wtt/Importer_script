# Register daily Task Scheduler job for eSSL -> Excel export (pywinauto).
#
# Default: every day at 11:50 (11:50 AM) and 23:50 (11:50 PM)
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -Unregister
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -DailyTimes 11:50,23:50

param(
    [string]$TaskName = "DDO-eSSL-Export",
    [string[]]$DailyTimes = @("11:50", "23:50"),
    # Legacy single-time alias (if set, overrides DailyTimes to one value)
    [string]$At = "",
    [switch]$Unregister,
    [switch]$SkipSync
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$Script = Join-Path $Root "scripts\essl_export.py"

function Find-Python {
    foreach ($dir in @(
        "${env:ProgramFiles}\Python313",
        "${env:ProgramFiles}\Python312",
        "${env:ProgramFiles}\Python311",
        "${env:ProgramFiles}\Python310",
        "${env:LOCALAPPDATA}\Programs\Python\Python313",
        "${env:LOCALAPPDATA}\Programs\Python\Python312",
        "${env:LOCALAPPDATA}\Programs\Python\Python311",
        "${env:LOCALAPPDATA}\Programs\Python\Python310"
    )) {
        $exe = Join-Path $dir "python.exe"
        if (Test-Path $exe) { return $exe }
    }
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notmatch 'WindowsApps') { return $cmd.Source }
    throw "Python 3 was not found. Re-run Setup.bat."
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Unregistered $TaskName"
    exit 0
}

if (-not (Test-Path $Script)) {
    throw "essl_export.py not found in $Root\scripts"
}

if ($At) {
    $DailyTimes = @($At)
}

# Export only. DDO-Attendance-Importer uploads separately and waits for this task if needed.
$Python = Find-Python
$argsList = "`"$Script`""
if ($SkipSync) { $argsList += " --skip-sync" }

$action = New-ScheduledTaskAction `
    -Execute $Python `
    -Argument $argsList `
    -WorkingDirectory $Root

$trigger = foreach ($time in $DailyTimes) {
    New-ScheduledTaskTrigger -Daily -At $time
}

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

$timesLabel = $DailyTimes -join " and "
$description = "Export attendance Excel from eSSL eTimeTrackLite into D:\Attendance daily at $timesLabel — 11:50 AM and 11:50 PM (catches up after a missed start)."

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description $description `
    -Force | Out-Null

Write-Host "Registered $TaskName daily at $timesLabel (export only)"
Write-Host "  action: $Python $argsList"
Write-Host "  dates: From=To=today (local PC date); override with ESSL_REPORT_DATE=YYYY-MM-DD"
Write-Host "Run once now:"
Write-Host "  python scripts\essl_export.py --login-only"
Write-Host "  python scripts\essl_export.py --select-only"
Write-Host "  python scripts\essl_export.py --skip-sync"
Write-Host "  python scripts\essl_export.py"
