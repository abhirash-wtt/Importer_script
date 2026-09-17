# Register a daily Task Scheduler job for eSSL -> Excel export (pywinauto).
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -Unregister
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_essl_export_scheduler.ps1 -At "07:30"

param(
    [string]$TaskName = "DDO-eSSL-Export",
    [string]$At = "07:30",
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

$argsList = @($Script)
if ($SkipSync) { $argsList += "--skip-sync" }

$action = New-ScheduledTaskAction `
    -Execute $Python `
    -Argument ($argsList -join " ") `
    -WorkingDirectory $Root

$trigger = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Export Daily Attendance Excel from eSSL eTimeTrackLite into D:\Attendance" `
    -Force | Out-Null

Write-Host "Registered $TaskName daily at $At"
Write-Host "  python: $Python"
Write-Host "  script: $Script"
Write-Host "Run once now:"
Write-Host "  python scripts\essl_export.py --login-only"
Write-Host "  python scripts\essl_export.py --select-only"
Write-Host "  python scripts\essl_export.py --skip-sync"
Write-Host "  python scripts\essl_export.py"
