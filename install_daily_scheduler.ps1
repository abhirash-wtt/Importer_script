# Registers a Windows Task Scheduler job that runs scheduler.ps1.
#
# Twice daily at 01:00 and 13:00 local time (default):
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_daily_scheduler.ps1
#
# Office 10-minute cadence (from the first DailyTimes value, Mon-Sat):
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_daily_scheduler.ps1 -EveryMinutes 10
#
# Remove the task:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_daily_scheduler.ps1 -Unregister

param(
    [string]$TaskName = "DDO-Attendance-Importer",
    [string[]]$DailyTimes = @("01:00", "13:00"),
    [int]$EveryMinutes = 0,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Scheduler = Join-Path $Root "scheduler.ps1"
if (-not (Test-Path $Scheduler)) {
    throw "scheduler.ps1 was not found in $Root"
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    exit 0
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Scheduler`"" -WorkingDirectory $Root
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

if ($EveryMinutes -gt 0) {
    $startTime = $DailyTimes[0]
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday, Saturday -At $startTime
    $trigger.RepetitionInterval = [TimeSpan]::FromMinutes($EveryMinutes)
    $trigger.RepetitionDuration = [TimeSpan]::FromHours(12)
    $description = "Import Agra, Noida, and Hyderabad attendance every $EveryMinutes minutes from $startTime."
} else {
    $trigger = foreach ($time in $DailyTimes) {
        New-ScheduledTaskTrigger -Daily -At $time
    }
    $timesLabel = $DailyTimes -join " and "
    $description = "Import Agra, Noida, and Hyderabad attendance daily at $timesLabel."
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description $description -Force | Out-Null
Write-Host "Registered '$TaskName'."
Write-Host $description
Write-Host "Drop .xls files in:"
Write-Host "  $Root\inbox\AGRA"
Write-Host "  $Root\inbox\NOIDA"
Write-Host "  $Root\inbox\HYD"
Write-Host "Test now: powershell -NoProfile -ExecutionPolicy Bypass -File `"$Scheduler`""
