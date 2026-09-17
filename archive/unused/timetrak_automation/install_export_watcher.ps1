# Optional: moves Excel from ESS_EXPORT_WATCH_DIR into ATTENDANCE_DIR (flat).
# Not needed for the simple HR flow — drop files directly into ATTENDANCE_DIR.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_export_watcher.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install_export_watcher.ps1 -Unregister

param(
    [string]$TaskName = "DDO-Attendance-Export-Watcher",
    [int]$EveryMinutes = 5,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Watcher = Join-Path $Root "watch_exports.ps1"
if (-not (Test-Path $Watcher)) {
    throw "watch_exports.ps1 was not found in $Root"
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    exit 0
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Watcher`"" `
    -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes) -RepetitionDuration ([TimeSpan]::MaxValue)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Move eSSL TimeTrak Excel exports into the DDO attendance drop folder." -Force | Out-Null

Write-Host "Registered '$TaskName' every $EveryMinutes minute(s)."
Write-Host "Only needed if TimeTrak saves outside ATTENDANCE_DIR."
Write-Host "For simple HR drops: leave ESS_EXPORT_WATCH_DIR blank and put Excel in ATTENDANCE_DIR."
Write-Host "Test now: powershell -NoProfile -ExecutionPolicy Bypass -File `"$Watcher`""
