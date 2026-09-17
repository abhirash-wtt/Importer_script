# Optional helper: if TimeTrak saves Excel somewhere else, move it into the
# single office drop folder (ATTENDANCE_DIR). For the simple HR flow you do NOT
# need this — just drop .xls/.xlsx straight into ATTENDANCE_DIR.
#
# Manual run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\watch_exports.ps1
#
# Continuous watch:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\watch_exports.ps1 -Continuous

param(
    [switch]$Continuous,
    [int]$PollSeconds = 10
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Get-DotEnvValue {
    param([string]$Key, [string]$Default = "")
    $envFile = Join-Path $Root ".env"
    if (-not (Test-Path $envFile)) { return $Default }
    foreach ($raw in Get-Content $envFile) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith("#") -or -not $line.Contains("=")) { continue }
        $eq = $line.IndexOf("=")
        $name = $line.Substring(0, $eq).Trim()
        if ($name -ne $Key) { continue }
        $value = $line.Substring($eq + 1).Trim().Trim('"').Trim("'")
        return $value
    }
    return $Default
}

function Get-LocationCode {
    $code = (Get-DotEnvValue "LOCATION_CODE" "").Trim().ToUpperInvariant()
    if ($code -eq "HYDERABAD") { $code = "HYD" }
    if ($code -notin @("AGRA", "NOIDA", "HYD")) {
        throw "Set LOCATION_CODE in .env to AGRA, NOIDA, or HYD."
    }
    return $code
}

function Get-DropDir {
    $attendance = (Get-DotEnvValue "ATTENDANCE_DIR" "D:\Attendance").Trim()
    if (-not $attendance) {
        throw "Set ATTENDANCE_DIR in .env (single folder where HR drops Excel)."
    }
    New-Item -ItemType Directory -Force -Path $attendance | Out-Null
    return $attendance
}

function Get-WatchDir {
    $configured = (Get-DotEnvValue "ESS_EXPORT_WATCH_DIR" "").Trim()
    if ($configured) { return $configured }
    return ""
}

function Test-FileReady {
    param([string]$Path, [int]$StableSeconds)
    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    } catch {
        return $false
    }
    $age = (Get-Date) - $item.LastWriteTime
    if ($age.TotalSeconds -lt $StableSeconds) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, "Open", "Read", "None")
        $stream.Close()
        return $true
    } catch {
        return $false
    }
}

function Move-NewExports {
    param(
        [string]$WatchDir,
        [string]$TargetDir,
        [int]$StableSeconds
    )

    if (-not $WatchDir) {
        Write-Host "ESS_EXPORT_WATCH_DIR is blank — nothing to move. Drop Excel directly into $TargetDir"
        return 0
    }
    if (-not (Test-Path -LiteralPath $WatchDir)) {
        Write-Host "Watch folder not found: $WatchDir"
        return 0
    }

    $watchFull = (Resolve-Path -LiteralPath $WatchDir).Path
    $targetFull = (Resolve-Path -LiteralPath $TargetDir).Path
    if ($watchFull -eq $targetFull) {
        Write-Host "Watch folder is the same as drop folder — importer will pick files up directly."
        return 0
    }

    $moved = 0
    $files = Get-ChildItem -LiteralPath $WatchDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(xls|xlsx)$' -and -not $_.Name.StartsWith("~$") }

    foreach ($file in $files) {
        if (-not (Test-FileReady -Path $file.FullName -StableSeconds $StableSeconds)) {
            continue
        }

        $destination = Join-Path $TargetDir $file.Name
        if (Test-Path -LiteralPath $destination) {
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $destination = Join-Path $TargetDir ("{0}_{1}{2}" -f $file.BaseName, $stamp, $file.Extension)
        }

        Move-Item -LiteralPath $file.FullName -Destination $destination -Force
        Write-Host "Moved $($file.Name) -> $destination"
        $moved += 1
    }
    return $moved
}

$location = Get-LocationCode
$dropDir = Get-DropDir
$watchDir = Get-WatchDir
$stable = [int](Get-DotEnvValue "INBOX_STABLE_SECONDS" "15")
if ($stable -lt 1) { $stable = 15 }

Write-Host "Location: $location (from .env)"
Write-Host "Drop:     $dropDir"
Write-Host "Watch:    $(if ($watchDir) { $watchDir } else { '(none — drop Excel into Drop folder)' })"
Write-Host "Stable:   ${stable}s"

do {
    $count = Move-NewExports -WatchDir $watchDir -TargetDir $dropDir -StableSeconds $stable
    if (-not $Continuous) {
        Write-Host "Moved $count file(s)."
        break
    }
    Start-Sleep -Seconds $PollSeconds
} while ($true)
