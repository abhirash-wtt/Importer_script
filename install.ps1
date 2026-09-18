# One-click office setup for DDO++ Attendance Agent (Python only).
#
# Double-click Install.bat, or run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
#
# Registers both Task Scheduler jobs by default (Mon-Fri 13:00):
#   DDO-Attendance-Importer
#   DDO-eSSL-Export
#
# Optional:
#   .\install.ps1 -SkipSchedulers
#   .\install.ps1 -SkipEnvEdit

param(
    [switch]$SkipSchedulers,
    [switch]$SkipEnvEdit,
    # Kept for compatibility; schedulers register by default now
    [switch]$RegisterSchedulers
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Find-Python {
    $candidates = @()
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
        if (Test-Path $exe) { $candidates += $exe }
    }
    foreach ($name in @("python", "python3")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -notmatch 'WindowsApps' -and $cmd.Source -notmatch 'python3\.1\dt') {
            $candidates += $cmd.Source
        }
    }
    foreach ($exe in ($candidates | Select-Object -Unique)) {
        try {
            $ver = & $exe --version 2>&1 | Out-String
            if ($LASTEXITCODE -eq 0 -and $ver -match 'Python 3') {
                return $exe
            }
        } catch { }
    }
    return $null
}

Write-Host "DDO++ Attendance Agent - install"
Write-Host "Repo: $Root"

Write-Step "Checking Python"
$Python = Find-Python
if (-not $Python) {
    throw "Python 3.10+ not found. Install from https://www.python.org/downloads/ (check Add python.exe to PATH), then re-run Install.bat."
}
Write-Host "Using: $Python"
& $Python --version

Write-Step "Creating runtime folders"
foreach ($dir in @("inbox", "processed", "failed", "logs", "output")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $dir) | Out-Null
}

Write-Step "Preparing .env"
$envPath = Join-Path $Root ".env"
$examplePath = Join-Path $Root ".env.example"
if (-not (Test-Path $examplePath)) {
    throw ".env.example is missing from the repo."
}
$createdEnv = $false
if (-not (Test-Path $envPath)) {
    Copy-Item $examplePath $envPath
    $createdEnv = $true
    Write-Host "Created .env from .env.example"
} else {
    Write-Host ".env already exists (left unchanged)"
}

Write-Step "Installing Python packages"
& $Python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed" }
& $Python -m pip install -r (Join-Path $Root "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "pip install -r requirements.txt failed" }

Write-Step "Creating attendance drop folder"
$attendanceDir = "D:\Attendance"
$envLines = Get-Content $envPath -ErrorAction SilentlyContinue
foreach ($line in $envLines) {
    if ($line -match '^\s*ATTENDANCE_DIR\s*=\s*(.+)\s*$') {
        $attendanceDir = $Matches[1].Trim().Trim('"').Trim("'")
    }
}
New-Item -ItemType Directory -Force -Path $attendanceDir | Out-Null
Write-Host "Drop folder: $attendanceDir"

$tokenMissing = $true
$locationOk = $false
foreach ($line in (Get-Content $envPath)) {
    if ($line -match '^\s*ATTENDANCE_INTEGRATION_TOKEN\s*=\s*(.+)\s*$' -and $Matches[1].Trim()) {
        $tokenMissing = $false
    }
    if ($line -match '^\s*DDO_API_TOKEN\s*=\s*(.+)\s*$' -and $Matches[1].Trim()) {
        $tokenMissing = $false
    }
    if ($line -match '^\s*LOCATION_CODE\s*=\s*(AGRA|NOIDA|HYD)\s*$') {
        $locationOk = $true
    }
}

if (-not $SkipEnvEdit -and ($createdEnv -or $tokenMissing -or -not $locationOk)) {
    Write-Step "Edit .env (set LOCATION_CODE and ATTENDANCE_INTEGRATION_TOKEN)"
    Write-Host "Save and close Notepad when done."
    Start-Process -FilePath "notepad.exe" -ArgumentList $envPath -Wait
}

if (-not $SkipSchedulers) {
    Write-Step "Registering Windows Task Scheduler jobs (Mon-Fri 13:00)"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_daily_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Failed to register DDO-Attendance-Importer" }
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_essl_export_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Failed to register DDO-eSSL-Export" }
    Write-Host "Registered: DDO-Attendance-Importer + DDO-eSSL-Export"
} else {
    Write-Host "Skipped Task Scheduler registration (-SkipSchedulers)"
}

Write-Host ""
Write-Host "Install complete." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Confirm .env has LOCATION_CODE and ATTENDANCE_INTEGRATION_TOKEN"
Write-Host "  2. Drop Excel into: $attendanceDir"
Write-Host "  3. Keep Windows logged on and unlocked for eSSL export (monitor off is OK)"
Write-Host "  4. Test import:"
Write-Host ('       "{0}" scripts\importer_script.py --inbox' -f $Python)
Write-Host "  5. Test eSSL export (monthly basic):"
Write-Host ('       "{0}" scripts\essl_export.py --skip-sync' -f $Python)
if ($SkipSchedulers) {
    Write-Host "  6. Register schedulers later:"
    Write-Host "       powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1"
}
Write-Host ""
Write-Host "Full guide: README.md"
