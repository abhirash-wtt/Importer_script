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
$Chain = Join-Path $Root "scripts\run_now.ps1"

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Unregistered $TaskName"
    exit 0
}

if (-not (Test-Path $Script)) {
    throw "essl_export.py not found in $Root\scripts"
}
if (-not (Test-Path $Chain)) {
    throw "run_now.ps1 not found in $Root\scripts"
}

if ($At) {
    $DailyTimes = @($At)
}

# Export then import in one task, so a catch-up run after a missed start
# (power cut, PC off) uploads straight after the export instead of racing it.
$argsList = "-NoProfile -ExecutionPolicy Bypass -File `"$Chain`" -Scheduled"
if ($SkipSync) { $argsList += " -SkipSync" }

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
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
$description = "Export attendance from eSSL into D:\Attendance, then upload to DDO++, daily at $timesLabel (catches up after a missed start)."

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description $description `
    -Force | Out-Null

Write-Host "Registered $TaskName daily at $timesLabel (export, then import)"
Write-Host "  action: powershell.exe $argsList"
Write-Host "  dates: From=To=today (local PC date); override with ESSL_REPORT_DATE=YYYY-MM-DD"
Write-Host "Run once now:"
Write-Host "  python scripts\essl_export.py --login-only"
Write-Host "  python scripts\essl_export.py --select-only"
Write-Host "  python scripts\essl_export.py --skip-sync"
Write-Host "  python scripts\essl_export.py"
