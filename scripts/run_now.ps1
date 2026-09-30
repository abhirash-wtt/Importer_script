# Manual "Run attendance now" button for HR (use when the scheduled run was missed).
#
# Step 1: eSSL export (scripts\essl_export.py) -> D:\Attendance\{Month} {Location}.xls
# Step 2: importer (scripts\scheduler.ps1) -> DDO++ API, only if step 1 succeeded
#
# Started from the "Run Attendance Now" Desktop / Start Menu shortcut created by Setup.bat.

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
Set-Location $Root
$Host.UI.RawUI.WindowTitle = "DDO++ Attendance - Run Now"

function Write-Banner([string]$Text, [string]$Color = "Cyan") {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor $Color
    Write-Host " $Text" -ForegroundColor $Color
    Write-Host "============================================================" -ForegroundColor $Color
}

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
    throw "Python 3 was not found. Ask IT to re-run Setup.bat."
}

function Test-TaskRunning([string]$TaskName) {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    return ($task -and $task.State -eq "Running")
}

function Wait-AndExit([int]$Code) {
    Write-Host ""
    Read-Host "Press Enter to close this window"
    exit $Code
}

Write-Banner "DDO++ Attendance - Run Now"
Write-Host " Step 1: Export attendance from eSSL"
Write-Host " Step 2: Upload it to DDO++"
Write-Host ""
Write-Host " Do not use the mouse/keyboard while eSSL is being automated." -ForegroundColor Yellow

foreach ($taskName in @("DDO-eSSL-Export", "DDO-Attendance-Importer")) {
    if (Test-TaskRunning $taskName) {
        Write-Banner "The scheduled run ($taskName) is already running. Try again in 15 minutes." "Yellow"
        Wait-AndExit 2
    }
}

try {
    $python = Find-Python
} catch {
    Write-Banner $_.Exception.Message "Red"
    Wait-AndExit 1
}

Write-Banner "Step 1/2: eSSL export (this can take several minutes)"
& $python (Join-Path $Root "scripts\essl_export.py")
$exportCode = $LASTEXITCODE
if ($exportCode -ne 0) {
    Write-Banner "Step 1 FAILED (exit code $exportCode). Upload was NOT started." "Red"
    Write-Host " Log: $(Join-Path $Root 'logs\essl_export.log')"
    Write-Host " Please inform the dev team."
    Wait-AndExit $exportCode
}

Write-Banner "Step 2/2: Uploading attendance to DDO++"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\scheduler.ps1")
$importCode = $LASTEXITCODE
if ($importCode -ne 0) {
    Write-Banner "Step 2 FAILED (exit code $importCode)." "Red"
    Write-Host " Log: $(Join-Path $Root 'output\importer.log')"
    Write-Host " Please inform the dev team."
    Wait-AndExit $importCode
}

Write-Banner "Done. Attendance exported and uploaded successfully." "Green"
Wait-AndExit 0
