param(
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$WindowTitle = "eSSL eTimeTrackLite"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WinCapture {
  [DllImport("user32.dll")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@

$dir = Split-Path -Parent $OutFile
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

$hwnd = [IntPtr]::Zero
Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$WindowTitle*" } | ForEach-Object {
    if ($hwnd -eq [IntPtr]::Zero) { $hwnd = $_.MainWindowHandle }
}

if ($hwnd -ne [IntPtr]::Zero) {
    [WinCapture]::ShowWindow($hwnd, 9) | Out-Null
    [WinCapture]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 400
    $rect = New-Object WinCapture+RECT
    if ([WinCapture]::GetWindowRect($hwnd, [ref]$rect)) {
        $w = [Math]::Max(1, $rect.Right - $rect.Left)
        $h = [Math]::Max(1, $rect.Bottom - $rect.Top)
        $bmp = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size($w, $h)))
        $bmp.Save($OutFile)
        $g.Dispose()
        $bmp.Dispose()
        Write-Output $OutFile
        exit 0
    }
}

# Fallback: full primary screen
$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size)
$bmp.Save($OutFile)
$g.Dispose()
$bmp.Dispose()
Write-Output $OutFile
