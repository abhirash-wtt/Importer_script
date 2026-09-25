# Register daily Task Scheduler job for eSSL -> Excel export (pywinauto).
#
# Default: every day at 01:00 (1 AM) and 13:00 (1 PM)
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -Unregister
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -DailyTimes 01:00,13:00

param(
    [string]$TaskName = "DDO-eSSL-Export",
    [string[]]$DailyTimes = @("01:00", "13:00"),
    # Legacy single-time alias (if set, overrides DailyTimes to one value)
    [string]$At = "",
    [switch]$Unregister,
    [switch]$SkipSync
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$Script = Join-Path $Root "scripts\essl_export.py"
$Python = (Get-Command python -ErrorAction Stop).Source

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

$argsList = @($Script)
if ($SkipSync) { $argsList += "--skip-sync" }

$action = New-ScheduledTaskAction `
    -Execute $Python `
    -Argument ($argsList -join " ") `
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
$description = "Export attendance Excel from eSSL eTimeTrackLite into D:\Attendance daily at $timesLabel (From=To=today)."

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description $description `
    -Force | Out-Null

Write-Host "Registered $TaskName daily at $timesLabel"
Write-Host "  python: $Python"
Write-Host "  script: $Script"
Write-Host "  dates: From=To=today (local PC date); override with ESSL_REPORT_DATE=YYYY-MM-DD"
Write-Host "Run once now:"
Write-Host "  python scripts\essl_export.py --login-only"
Write-Host "  python scripts\essl_export.py --select-only"
Write-Host "  python scripts\essl_export.py --skip-sync"
Write-Host "  python scripts\essl_export.py"
