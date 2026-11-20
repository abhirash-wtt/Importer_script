# Registers a Windows Task Scheduler job that runs scripts\scheduler.ps1.
#
# Default: every day at 12:00 (noon) and 00:00 (midnight) — 10 minutes after eSSL export at 11:50 / 23:50
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1
#
# Custom times:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1 -DailyTimes 12:00,00:00
#
# Remove the task:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_daily_scheduler.ps1 -Unregister

param(
    [string]$TaskName = "DDO-Attendance-Importer",
    [string[]]$DailyTimes = @("12:00", "00:00"),
    [int]$EveryMinutes = 0,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
$Scheduler = Join-Path $Root "scripts\scheduler.ps1"
if (-not (Test-Path $Scheduler)) {
    throw "scheduler.ps1 was not found in $Root\scripts"
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    exit 0
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Scheduler`"" -WorkingDirectory $Root
# StartWhenAvailable: catch up after a missed start (power cut). scheduler.ps1 waits for
# DDO-eSSL-Export first if it is running or has not yet run for its slot.
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

if ($EveryMinutes -gt 0) {
    $startTime = $DailyTimes[0]
    $trigger = New-ScheduledTaskTrigger -Daily -At $startTime
    $trigger.RepetitionInterval = [TimeSpan]::FromMinutes($EveryMinutes)
    $trigger.RepetitionDuration = [TimeSpan]::FromHours(12)
    $description = "Import attendance every $EveryMinutes minutes from $startTime (daily)."
} else {
    $trigger = foreach ($time in $DailyTimes) {
        New-ScheduledTaskTrigger -Daily -At $time
    }
    $timesLabel = $DailyTimes -join " and "
    $description = "Import attendance daily at $timesLabel (after eSSL export)."
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description $description -Force | Out-Null
Write-Host "Registered '$TaskName'."
Write-Host $description
Write-Host "Drop Excel files directly into:"
Write-Host "  $((Get-Content (Join-Path $Root '.env') -ErrorAction SilentlyContinue | Where-Object { $_ -match '^ATTENDANCE_DIR=' }) -replace '^ATTENDANCE_DIR=','' -replace '\"','')"
Write-Host "  (default D:\Attendance - set ATTENDANCE_DIR in .env)"
Write-Host "Location for this PC comes from LOCATION_CODE in .env (no subfolders needed)."
Write-Host "Test now: powershell -NoProfile -ExecutionPolicy Bypass -File `"$Scheduler`""
Write-Host ""
Write-Host "After each run check:"
Write-Host "  $Root\output\last_run.txt"
Write-Host "  $Root\output\scheduler_runs.log"
Write-Host "  $Root\output\importer.log"
Write-Host "To enable Task Scheduler History (Admin PowerShell):"
Write-Host '  wevtutil set-log Microsoft-Windows-TaskScheduler/Operational /enabled:true'
