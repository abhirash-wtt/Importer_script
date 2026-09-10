# DDO++ attendance importer — office wrapper.
# Scans inbox\AGRA, inbox\NOIDA, and inbox\HYD, then POSTs each report.
#
# Manual run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scheduler.ps1

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Find-Python {
    $candidates = @()
    $localPython = Join-Path $env:LOCALAPPDATA "Programs\Python"
    if (Test-Path $localPython) {
        Get-ChildItem $localPython -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            ForEach-Object {
                $exe = Join-Path $_.FullName "python.exe"
                if (Test-Path $exe) {
                    $candidates += @{ File = $exe; Args = @() }
                }
            }
    }
    $pyLauncher = Join-Path $env:LOCALAPPDATA "Programs\Python\Launcher\py.exe"
    if (Test-Path $pyLauncher) {
        $candidates += @{ File = $pyLauncher; Args = @("-3") }
    }
    $candidates += @(
        @{ File = "py"; Args = @("-3") },
        @{ File = "python"; Args = @() },
        @{ File = "python3"; Args = @() }
    )
    foreach ($item in $candidates) {
        if ($item.File -notmatch '[\\/]' -and -not (Get-Command $item.File -ErrorAction SilentlyContinue)) {
            continue
        }
        if ($item.File -match '[\\/]' -and -not (Test-Path $item.File)) {
            continue
        }
        try {
            $versionArgs = @($item.Args) + @("--version")
            $null = & $item.File @versionArgs 2>$null
            if ($LASTEXITCODE -eq 0) {
                return $item
            }
        } catch {
            continue
        }
    }
    throw "Python 3 was not found. Install Python 3 or add python.exe to PATH."
}

$python = Find-Python
$importer = Join-Path $Root "importer_script.py"
if (-not (Test-Path $importer)) {
    throw "importer_script.py was not found in $Root"
}

$pythonArgs = $python.Args + @($importer, "--inbox")
& $python.File @pythonArgs
$exitCode = $LASTEXITCODE

$stampDir = Join-Path $Root "output"
New-Item -ItemType Directory -Force -Path $stampDir | Out-Null
$stamp = Join-Path $stampDir "last_run.txt"
$finished = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Set-Content -Path $stamp -Value "finished_at=$finished`r`nexit_code=$exitCode" -Encoding UTF8

exit $exitCode
