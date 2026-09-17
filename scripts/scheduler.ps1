# DDO++ attendance importer — office wrapper.
# Scans ATTENDANCE_DIR (single drop folder), converts Excel → JSON, POSTs to DDO++.
#
# Manual run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\scheduler.ps1
#
# Every run appends to output\scheduler_runs.log (works even when Task Scheduler History is off).

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $ScriptDir
Set-Location $Root

$stampDir = Join-Path $Root "output"
New-Item -ItemType Directory -Force -Path $stampDir | Out-Null
$runLog = Join-Path $stampDir "scheduler_runs.log"
$started = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

function Write-RunLog([string]$Message) {
    Add-Content -Path $runLog -Value $Message -Encoding UTF8
}

Write-RunLog "===== START $started pid=$PID ====="

try {
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
    $importer = Join-Path $Root "scripts\importer_script.py"
    if (-not (Test-Path $importer)) {
        throw "importer_script.py was not found in $Root\scripts"
    }

    $pythonArgs = $python.Args + @($importer, "--inbox")
    $importOut = & $python.File @pythonArgs 2>&1 | Out-String
    $exitCode = $LASTEXITCODE
    if ($importOut.Trim()) {
        Write-RunLog ($importOut.TrimEnd())
    }
} catch {
    $exitCode = 1
    Write-RunLog ("ERROR: " + $_.Exception.Message)
}

$finished = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$stamp = Join-Path $stampDir "last_run.txt"
Set-Content -Path $stamp -Value "finished_at=$finished`r`nexit_code=$exitCode" -Encoding UTF8
Write-RunLog "===== END $finished exit_code=$exitCode ====="
Write-RunLog ""

exit $exitCode
