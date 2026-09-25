# DDO++ Attendance Agent - full office install package (Python only).
#
# Double-click Setup.bat, or:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
#
# What it does:
#   1. Finds or installs Python 3.10+ (winget)
#   2. Creates runtime folders
#   3. Creates .env from README.md (```env-example)
#   4. pip installs packages listed in README.md (```pip-requirements)
#   5. Creates ATTENDANCE_DIR drop folder
#   6. Opens Notepad for LOCATION_CODE + token when needed
#   7. Registers Task Scheduler jobs (eSSL 01:00+13:00, importer 01:10+13:10 daily)
#   8. Adds Start Menu shortcuts
#
# Options:
#   .\install.ps1 -SkipSchedulers
#   .\install.ps1 -SkipEnvEdit
#   .\install.ps1 -SkipPythonInstall
#   .\install.ps1 -Uninstall

param(
    [switch]$SkipSchedulers,
    [switch]$SkipEnvEdit,
    [switch]$SkipPythonInstall,
    [switch]$Uninstall,
    # Kept for compatibility; schedulers register by default now
    [switch]$RegisterSchedulers
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$ProductName = "DDO++ Attendance Agent"
$StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$ProductName"

function Write-Banner {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " $ProductName - Setup" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Folder: $Root"
}

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
    foreach ($name in @("python", "python3", "py")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        if ($cmd.Source -match 'WindowsApps') { continue }
        if ($cmd.Source -match 'python3\.1\dt') { continue }
        if ($name -eq "py") {
            try {
                $resolved = & $cmd.Source -3 -c "import sys; print(sys.executable)" 2>$null
                if ($LASTEXITCODE -eq 0 -and $resolved) { $candidates += $resolved.Trim() }
            } catch { }
            continue
        }
        $candidates += $cmd.Source
    }
    foreach ($exe in ($candidates | Select-Object -Unique)) {
        try {
            $ver = & $exe --version 2>&1 | Out-String
            if ($LASTEXITCODE -eq 0 -and $ver -match 'Python 3\.(\d+)' -and [int]$Matches[1] -ge 10) {
                return $exe
            }
        } catch { }
    }
    return $null
}

function Refresh-PathFromMachine {
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machine;$user"
}

function Install-PythonWithWinget {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Python 3.10+ not found and winget is unavailable. Install Python from https://www.python.org/downloads/ (check Add python.exe to PATH), then re-run Setup.bat."
    }
    Write-Host "Installing Python 3.12 with winget (may take a few minutes)..."
    & winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "winget failed to install Python. Install manually from https://www.python.org/downloads/ then re-run Setup.bat."
    }
    Refresh-PathFromMachine
    Start-Sleep -Seconds 2
}

function Get-PackagesFromReadme {
    $readmeText = Get-ReadmeText
    if ($readmeText -notmatch '(?s)```pip-requirements\r?\n(.*?)```') {
        throw "Could not find ```pip-requirements block in README.md"
    }
    $packages = @()
    foreach ($line in ($Matches[1] -split '\r?\n')) {
        $pkg = $line.Trim()
        if ($pkg -and -not $pkg.StartsWith("#")) {
            $packages += $pkg
        }
    }
    if ($packages.Count -eq 0) {
        throw "pip-requirements block in README.md is empty"
    }
    return $packages
}

function Get-ReadmeText {
    $readmePath = Join-Path $Root "README.md"
    if (-not (Test-Path $readmePath)) {
        throw "README.md is missing from the repo."
    }
    return Get-Content -Path $readmePath -Raw
}

function Get-EnvExampleFromReadme {
    $readmeText = Get-ReadmeText
    if ($readmeText -notmatch '(?s)```env-example\r?\n(.*?)```') {
        throw "Could not find ```env-example block in README.md"
    }
    $body = $Matches[1] -replace '(\r?\n)+$', ''
    if (-not $body.Trim()) {
        throw "env-example block in README.md is empty"
    }
    return ($body + "`r`n")
}

function Get-AttendanceDir([string]$EnvPath) {
    $attendanceDir = "D:\Attendance"
    foreach ($line in (Get-Content $EnvPath -ErrorAction SilentlyContinue)) {
        if ($line -match '^\s*ATTENDANCE_DIR\s*=\s*(.+)\s*$') {
            $attendanceDir = $Matches[1].Trim().Trim('"').Trim("'")
        }
    }
    return $attendanceDir
}

function New-Shortcut([string]$LinkPath, [string]$TargetPath, [string]$Arguments = "", [string]$WorkingDirectory = "", [string]$Description = "") {
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($LinkPath)
    $sc.TargetPath = $TargetPath
    if ($Arguments) { $sc.Arguments = $Arguments }
    if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
    if ($Description) { $sc.Description = $Description }
    $sc.Save()
}

function Install-StartMenuShortcuts([string]$Python, [string]$AttendanceDir) {
    New-Item -ItemType Directory -Force -Path $StartMenuDir | Out-Null
    New-Shortcut (Join-Path $StartMenuDir "Open attendance drop folder.lnk") "explorer.exe" $AttendanceDir "" "Open ATTENDANCE_DIR"
    New-Shortcut (Join-Path $StartMenuDir "Edit office .env.lnk") "notepad.exe" (Join-Path $Root ".env") $Root "Edit LOCATION_CODE and API token"
    New-Shortcut (Join-Path $StartMenuDir "Open agent folder.lnk") "explorer.exe" $Root "" "Open install folder"
    New-Shortcut (Join-Path $StartMenuDir "View last import log.lnk") "notepad.exe" (Join-Path $Root "output\last_run.txt") (Join-Path $Root "output") "Open last_run.txt"
    New-Shortcut (Join-Path $StartMenuDir "Run importer now.lnk") $Python (Join-Path $Root "scripts\importer_script.py") $Root "Run importer once"
    New-Shortcut (Join-Path $StartMenuDir "Uninstall schedulers.lnk") "powershell.exe" "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $Root 'install.ps1')`" -Uninstall" $Root "Remove Task Scheduler jobs and shortcuts"
    Write-Host "Start Menu: $StartMenuDir"
}

function Uninstall-Agent {
    Write-Banner
    Write-Step "Removing Task Scheduler jobs"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_daily_scheduler.ps1") -Unregister
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_essl_export_scheduler.ps1") -Unregister

    Write-Step "Removing Start Menu shortcuts"
    if (Test-Path $StartMenuDir) {
        Remove-Item -Recurse -Force $StartMenuDir
        Write-Host "Removed: $StartMenuDir"
    } else {
        Write-Host "No Start Menu folder found."
    }

    Write-Host ""
    Write-Host "Uninstall complete." -ForegroundColor Green
    Write-Host "Left in place: this folder, .env, Excel drop files, Python, and pip packages."
    return
}

if ($Uninstall) {
    Uninstall-Agent
    exit 0
}

Write-Banner

# --- 1. Python ---
Write-Step "1/8 Checking Python"
$Python = Find-Python
if (-not $Python) {
    if ($SkipPythonInstall) {
        throw "Python 3.10+ not found. Install from https://www.python.org/downloads/ then re-run Setup.bat."
    }
    Write-Host "Python not found. Attempting automatic install..."
    Install-PythonWithWinget
    $Python = Find-Python
    if (-not $Python) {
        throw "Python was installed but is not on PATH yet. Close this window, open a new one, and run Setup.bat again."
    }
}
Write-Host "Using: $Python"
& $Python --version

# --- 2. Folders ---
Write-Step "2/8 Creating runtime folders"
foreach ($dir in @("processed", "failed", "logs", "output")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $dir) | Out-Null
}

# --- 3. .env ---
Write-Step "3/8 Preparing .env"
$envPath = Join-Path $Root ".env"
$createdEnv = $false
if (-not (Test-Path $envPath)) {
    $envTemplate = Get-EnvExampleFromReadme
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($envPath, $envTemplate, $utf8NoBom)
    $createdEnv = $true
    Write-Host "Created .env from README.md env-example template"
} else {
    Write-Host ".env already exists (left unchanged)"
}

# --- 4. pip packages ---
Write-Step "4/8 Installing Python packages (from README.md)"
$packages = Get-PackagesFromReadme
Write-Host ("Packages: " + ($packages -join ", "))
& $Python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed" }
& $Python -m pip install @packages
if ($LASTEXITCODE -ne 0) { throw "pip install from README.md failed" }

Write-Step "5/8 Verifying packages"
& $Python -c "import xlrd, openpyxl, pywinauto, comtypes; print('ok')"
if ($LASTEXITCODE -ne 0) { throw "Package import check failed" }

# --- 6. Drop folder ---
Write-Step "6/8 Creating attendance drop folder"
$attendanceDir = Get-AttendanceDir $envPath
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
    $attendanceDir = Get-AttendanceDir $envPath
    New-Item -ItemType Directory -Force -Path $attendanceDir | Out-Null
}

# --- 7. Schedulers ---
if (-not $SkipSchedulers) {
    Write-Step "7/8 Registering Windows Task Scheduler jobs (eSSL 01:00+13:00, importer 01:10+13:10 daily)"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_daily_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Failed to register DDO-Attendance-Importer" }
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "scripts\install_essl_export_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Failed to register DDO-eSSL-Export" }
    Write-Host "Registered: DDO-eSSL-Export (01:00 & 13:00) + DDO-Attendance-Importer (01:10 & 13:10) daily"
} else {
    Write-Step "7/8 Skipped Task Scheduler registration (-SkipSchedulers)"
}

# --- 8. Shortcuts ---
Write-Step "8/8 Creating Start Menu shortcuts"
Install-StartMenuShortcuts -Python $Python -AttendanceDir $attendanceDir

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Setup complete." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Confirm .env has LOCATION_CODE and ATTENDANCE_INTEGRATION_TOKEN"
Write-Host "  2. Drop Excel into: $attendanceDir"
Write-Host "  3. Keep Windows logged on and unlocked for eSSL export (monitor off is OK)"
Write-Host "  4. Start Menu -> $ProductName for shortcuts"
Write-Host "  5. Test import:"
Write-Host ('       "{0}" scripts\importer_script.py' -f $Python)
Write-Host "  6. Test eSSL export (monthly basic):"
Write-Host ('       "{0}" scripts\essl_export.py --skip-sync' -f $Python)
if ($SkipSchedulers) {
    Write-Host "  7. Register schedulers later by running Setup.bat again"
}
Write-Host ""
Write-Host "Uninstall schedulers/shortcuts: Uninstall.bat"
Write-Host "Full guide: README.md"
