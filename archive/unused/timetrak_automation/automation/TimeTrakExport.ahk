; =============================================================================
; eSSL eTimeTrackLite 12.0 — full office flow (AutoHotkey v2)
;
; Confirmed office steps:
;   0) Login essl / essl
;   1) Utilities → Device Management
;      Select 4 devices (LGF OUT, LGF IN, UGF OUT, UGF IN 1) — skip USB
;      Start Download → wait until sync finishes
;   2) Attendance Reports → Daily Attendance Reports → Basic Report
;      Filter box opens → choose report type (Basic Report) → Generate
;   3) Report viewer → toolbar export (floppy) → Excel → Save As
;      Save into D:\Attendance\exports
; =============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
; Match mouse coords to screen pixels on scaled displays
try DllCall("SetProcessDPIAware")
SetTitleMatchMode 2
CoordMode "Mouse", "Client"
CoordMode "Pixel", "Window"

WINDOW_TITLE := "eSSL eTimeTrackLite"
LOGIN_TITLE := "eSSL eTimeTrackLite Login"
REPORT_FILTER_TITLE := "Daily Attendance Report"
DEVICE_MGMT_TITLE := "Device Management"
EXPORT_DIR := "D:\Attendance\exports"
EXPORT_NAME := Format("Daily_Attendance_Report_{}.xls", FormatTime(, "yyyyMMdd_HHmmss"))

LOGIN_NAME := "essl"
LOGIN_PASSWORD := "essl"

APP_EXE_CANDIDATES := [
    "C:\Program Files (x86)\essl\eTimeTrackLite\eTimeTrackLite.exe",
    "C:\Program Files (x86)\eSSL\eTimeTrackLite\eTimeTrackLite.exe",
    "C:\Program Files\essl\eTimeTrackLite\eTimeTrackLite.exe"
]

DOWNLOAD_TIMEOUT_SEC := 600
DOWNLOAD_POLL_SEC := 5
UNATTENDED := false
LOGIN_ONLY := false
SYNC_ONLY := false
SELECT_ONLY := false

if A_Args.Length >= 1 && (A_Args[1] = "--test" || A_Args[1] = "test") {
    DOWNLOAD_TIMEOUT_SEC := 60
    UNATTENDED := false
}
if A_Args.Length >= 1 && (A_Args[1] = "--login-only" || A_Args[1] = "login-only") {
    LOGIN_ONLY := true
}
if A_Args.Length >= 1 && (A_Args[1] = "--sync-only" || A_Args[1] = "sync-only") {
    SYNC_ONLY := true
    DOWNLOAD_TIMEOUT_SEC := 120
}
if A_Args.Length >= 1 && (A_Args[1] = "--select-only" || A_Args[1] = "select-only") {
    SELECT_ONLY := true
}
EnsureDir(EXPORT_DIR)
EnsureAppRunning()
HandleLogin()

WinActivate WINDOW_TITLE
if !WinWaitActive(WINDOW_TITLE, , 30) {
    Fail("Could not focus eTimeTrackLite after login.")
}
Sleep 800

if LOGIN_ONLY {
    MsgBox "Login OK. Main window is active.`nStop here (--login-only)."
    ExitApp 0
}

; --- Step 1: Utilities → Device Management → select 4 devices → Start Download ---
OpenDeviceManagement()
SelectFourDevices()

if SELECT_ONLY {
    ; Keep Device Management focused so the end screenshot is TimeTrak, not Cursor.
    WinActivate WINDOW_TITLE
    WinWaitActive WINDOW_TITLE, , 5
    Sleep 800
    SnapAutomationScreen("select-only-before-end")
    Sleep 300
    MsgBox "Device checkbox clicks done.`nLook at Device List checkboxes now.`nScreenshot saved in output\.`nClick OK when finished checking."
    ExitApp 0
}

ClickStartDownload()
WaitForDownloadComplete(DOWNLOAD_TIMEOUT_SEC, DOWNLOAD_POLL_SEC)

if SYNC_ONLY {
    MsgBox "Device sync wait finished (--sync-only).`nCheck Status / Logs Downloaded columns."
    ExitApp 0
}

; --- Step 2: Attendance Reports → Daily → Basic Report → Generate ---
OpenDailyBasicReport()
ChooseBasicReportAndGenerate()

; --- Step 3: Export Excel → Save As ---
ExportReportToExcel()
HandleSaveAs(EXPORT_DIR, EXPORT_NAME)

Sleep 3000
if !UNATTENDED {
    MsgBox "Export finished. Check:`n" EXPORT_DIR "\`n" EXPORT_NAME
}
ExitApp 0

; ---------------------------------------------------------------------------
EnsureAppRunning() {
    global WINDOW_TITLE, LOGIN_TITLE, APP_EXE_CANDIDATES
    if WinExist(WINDOW_TITLE) || WinExist(LOGIN_TITLE)
        return
    exe := ""
    for path in APP_EXE_CANDIDATES {
        if FileExist(path) {
            exe := path
            break
        }
    }
    if exe = "" {
        Fail("TimeTrak exe not found. Update APP_EXE_CANDIDATES.")
    }
    Run exe
    if !(WinWait(WINDOW_TITLE, , 45) || WinWait(LOGIN_TITLE, , 45)) {
        Fail("TimeTrak did not start.")
    }
}

HandleLogin() {
    global LOGIN_TITLE, LOGIN_NAME, LOGIN_PASSWORD, WINDOW_TITLE

    Sleep 800
    if !WinExist(LOGIN_TITLE) {
        if !WinWait(LOGIN_TITLE, , 12) {
            if WinExist(WINDOW_TITLE)
                return
            Fail("Login window did not appear.")
        }
    }

    loginHwnd := WinExist(LOGIN_TITLE)
    WinActivate "ahk_id " loginHwnd
    if !WinWaitActive("ahk_id " loginHwnd, , 10) {
        Fail("Could not focus the login window.")
    }
    Sleep 500

    filled := false
    try {
        ControlSetText LOGIN_NAME, "Edit1", "ahk_id " loginHwnd
        Sleep 200
        ControlSetText LOGIN_PASSWORD, "Edit2", "ahk_id " loginHwnd
        Sleep 200
        filled := true
    } catch {
        filled := false
    }

    if !filled {
        Click 160, 115
        Sleep 250
        Send "^a"
        Send LOGIN_NAME
        Sleep 200
        Send "{Tab}"
        Sleep 200
        Send "^a"
        Send LOGIN_PASSWORD
        Sleep 250
    }

    clicked := false
    try {
        ControlClick "Login", "ahk_id " loginHwnd
        clicked := true
    } catch {
        try {
            ControlClick "Button1", "ahk_id " loginHwnd
            clicked := true
        } catch {
            clicked := false
        }
    }
    if !clicked {
        Send "{Tab}"
        Sleep 150
        Send "{Enter}"
    }

    if !WinWaitClose("ahk_id " loginHwnd, , 25) {
        WinActivate "ahk_id " loginHwnd
        Sleep 300
        Click 160, 115
        Sleep 200
        Send "^a"
        Send LOGIN_NAME
        Send "{Tab}"
        Sleep 150
        Send "^a"
        Send LOGIN_PASSWORD
        Sleep 200
        Send "{Enter}"
        if !WinWaitClose(LOGIN_TITLE, , 20) {
            Fail("Login failed — dialog stayed open after essl/essl.")
        }
    }

    if !WinWait(WINDOW_TITLE, , 20) {
        Fail("Main window did not appear after login.")
    }
    Sleep 1200
    if WinExist(LOGIN_TITLE) {
        Fail("Login dialog is still visible after submit.")
    }
}

OpenDeviceManagement() {
    global WINDOW_TITLE, DEVICE_MGMT_TITLE
    WinActivate WINDOW_TITLE
    Sleep 400

    ; Close popups that block the device grid (Users List was opened by mistake via Alt+U)
    CloseBlockingDialogs()

    if InStr(WinGetTitle("A"), DEVICE_MGMT_TITLE) {
        return
    }

    ; Do NOT use Alt+U — that opens Admin → Users on this build.
    ; Open Utilities with Alt + Right arrows:
    ;   Admin → Masters → Utilities
    Send "{Alt}"
    Sleep 350
    Send "{Right}"   ; Masters
    Sleep 200
    Send "{Right}"   ; Utilities
    Sleep 350
    Send "{Down}"    ; open Utilities submenu / first item
    Sleep 300
    ; Device Management is usually near the top of Utilities
    Send "{Enter}"
    Sleep 1200

    CloseBlockingDialogs()

    if !InStr(WinGetTitle("A"), DEVICE_MGMT_TITLE) {
        ; Try next Utilities item
        Send "{Alt}"
        Sleep 300
        Send "{Right}"
        Sleep 150
        Send "{Right}"
        Sleep 300
        Send "{Down}"
        Sleep 200
        Send "{Down}"
        Sleep 200
        Send "{Enter}"
        Sleep 1200
    }

    CloseBlockingDialogs()
}

CloseBlockingDialogs() {
    global WINDOW_TITLE
    ; Windows that steal clicks away from Device List checkboxes
    blockers := ["Users List", "User List", "Employee", "Add User", "Edit User"]
    for title in blockers {
        if WinExist(title) {
            WinActivate title
            Sleep 200
            Send "{Esc}"
            Sleep 300
            if WinExist(title) {
                WinClose title
                Sleep 400
            }
        }
    }
    if WinExist(WINDOW_TITLE) {
        WinActivate WINDOW_TITLE
        Sleep 300
    }
}

SelectFourDevices() {
    global WINDOW_TITLE
    CoordMode "Mouse", "Client"
    WinActivate WINDOW_TITLE
    WinWaitActive WINDOW_TITLE, , 5
    Sleep 500
    CloseBlockingDialogs()
    Sleep 300

    DumpDeviceControls()

    grid := FindDeviceListGrid()
    checkX := 22
    headerCheckY := 136
    usbY := 158
    if grid != 0 {
        checkX := grid.x + 20
        headerCheckY := grid.y + 42
        usbY := headerCheckY + 22
    }
    FileAppend Format("select checkX={1} headerY={2} usbY={3} grid={4}`n", checkX, headerCheckY, usbY, grid = 0 ? "none" : grid.name), A_ScriptDir "\..\output\device_controls.txt"

    ; Select ALL via header checkbox, then uncheck USB only.
    ; (Clicking Device Name only highlights — must hit the checkbox square.)
    headerBtn := FindHeaderCheckButton()
    if headerBtn != 0 {
        ControlClick headerBtn.name, WINDOW_TITLE
        FileAppend Format("headerBtn={1}`n", headerBtn.name), A_ScriptDir "\..\output\device_controls.txt"
    } else {
        Click checkX, headerCheckY
    }
    Sleep 600
    Click checkX, usbY
    Sleep 500

    CoordMode "Mouse", "Window"
}

FindHeaderCheckButton() {
    global WINDOW_TITLE
    try {
        for ctrl in WinGetControls(WINDOW_TITLE) {
            if !InStr(ctrl, "BUTTON")
                continue
            ControlGetPos &cx, &cy, &cw, &ch, ctrl, WINDOW_TITLE
            if (cx >= 5 && cx <= 25 && cy >= 95 && cy <= 150 && cw <= 20 && ch <= 20)
                return { name: ctrl, x: cx, y: cy, w: cw, h: ch }
        }
    } catch {
        return 0
    }
    return 0
}

DumpDeviceControls() {
    global WINDOW_TITLE
    out := A_ScriptDir "\..\output\device_controls.txt"
    try FileDelete out
    lines := "=== Device window controls ===`n"
    lines .= "title=" WinGetTitle(WINDOW_TITLE) "`n"
    try {
        for ctrl in WinGetControls(WINDOW_TITLE) {
            ControlGetPos &cx, &cy, &cw, &ch, ctrl, WINDOW_TITLE
            lines .= Format("{1}`t x={2} y={3} w={4} h={5}`n", ctrl, cx, cy, cw, ch)
        }
    } catch as err {
        lines .= "error: " err.Message "`n"
    }
    FileAppend lines, out
}

FindDeviceListGrid() {
    global WINDOW_TITLE
    best := 0
    bestArea := 0
    try {
        for ctrl in WinGetControls(WINDOW_TITLE) {
            if !InStr(ctrl, "WindowsForms10.Window.8")
                continue
            ControlGetPos &cx, &cy, &cw, &ch, ctrl, WINDOW_TITLE
            ; Main Device List pane: near left, wide+tall, below options bar (~y 90+).
            if (cx <= 10 && cy >= 80 && cy < 130 && cw > 800 && ch > 400) {
                area := cw * ch
                if area > bestArea {
                    bestArea := area
                    best := { name: ctrl, x: cx, y: cy, w: cw, h: ch }
                }
            }
        }
    } catch {
        return 0
    }
    return best
}

SelectFourDevicesByKeyboard() {
    global WINDOW_TITLE
    CoordMode "Mouse", "Client"
    WinActivate WINDOW_TITLE
    Sleep 300

    ; Focus grid on Device Name of first row
    Click 120, 190
    Sleep 300
    Send "{Home}"
    Sleep 200

    ; Assume rows start mostly unchecked after a fresh screen open:
    ; skip USB, check next 4 with Space
    Loop 4 {
        Send "{Down}"
        Sleep 160
        Send "{Space}"
        Sleep 160
    }
    CoordMode "Mouse", "Window"
}

ClickStartDownload() {
    global WINDOW_TITLE
    WinActivate WINDOW_TITLE
    Sleep 300

    try {
        ControlClick "Start Download", WINDOW_TITLE
        Sleep 800
        return
    } catch {
    }

    Click 780, 118
    Sleep 800
}

WaitForDownloadComplete(timeoutSec, pollSec) {
    elapsed := 0
    while elapsed < timeoutSec {
        Sleep pollSec * 1000
        elapsed += pollSec
        if WinExist("Error") || WinExist("Attention") {
            Fail("Download reported an error dialog. Fix devices and retry.")
        }
    }
}

OpenDailyBasicReport() {
    global WINDOW_TITLE, REPORT_FILTER_TITLE
    WinActivate WINDOW_TITLE
    Sleep 500

    ; Admin → Masters → Utilities → Attendance Reports
    Send "{Alt}"
    Sleep 300
    Send "{Right}"
    Sleep 150
    Send "{Right}"
    Sleep 150
    Send "{Right}"
    Sleep 300
    Send "{Down}"
    Sleep 400
    Send "{Enter}"
    Sleep 500
    Send "{Enter}"
    Sleep 1500

    if !WinExist(REPORT_FILTER_TITLE) {
        WinActivate WINDOW_TITLE
        Send "!t"
        Sleep 600
        Send "{Enter}"
        Sleep 400
        Send "{Enter}"
        Sleep 1500
    }
}

ChooseBasicReportAndGenerate() {
    global REPORT_FILTER_TITLE
    if !WinWait(REPORT_FILTER_TITLE, , 20) {
        Fail("Daily Attendance Report filter window did not open.")
    }
    filterHwnd := WinExist(REPORT_FILTER_TITLE)
    WinActivate "ahk_id " filterHwnd
    Sleep 600

    ; Force Basic Report in the top-right dropdown
    try {
        ControlClick "ComboBox1", "ahk_id " filterHwnd
        Sleep 300
        Send "{Home}"
        Sleep 150
        Send "Basic Report"
        Sleep 200
        Send "{Enter}"
        Sleep 300
    } catch {
        Click 520, 70
        Sleep 300
        Send "{Home}"
        Sleep 150
        Send "{Enter}"
        Sleep 300
    }

    clicked := false
    try {
        ControlClick "Generate", "ahk_id " filterHwnd
        clicked := true
    } catch {
        try {
            ControlClick "Button1", "ahk_id " filterHwnd
            clicked := true
        } catch {
            clicked := false
        }
    }
    if !clicked {
        Send "!g"
        Sleep 300
        Click 560, 430
    }
    Sleep 3000
}

ExportReportToExcel() {
    global REPORT_FILTER_TITLE
    if !WinWait(REPORT_FILTER_TITLE, , 40) {
        Fail("Report viewer did not open after Generate.")
    }
    WinActivate REPORT_FILTER_TITLE
    Sleep 1000

    Click 280, 45
    Sleep 700
    Send "{Enter}"
    Sleep 1200
}

HandleSaveAs(exportDir, exportName) {
    if !(WinWait("Save As", , 25) || WinWait("Save", , 5)) {
        Fail("Save As dialog did not appear after Excel export.")
    }
    Sleep 400
    Send "!d"
    Sleep 250
    SendText exportDir
    Send "{Enter}"
    Sleep 500
    Send "!n"
    Sleep 200
    Send "^a"
    SendText exportName
    Sleep 200
    Send "{Enter}"
    Sleep 800
    if WinExist("Confirm Save As") || WinExist("Confirm") {
        Send "!y"
    }
}

EnsureDir(path) {
    if !DirExist(path)
        DirCreate path
}

SnapAutomationScreen(label) {
    outDir := A_ScriptDir "\..\output"
    EnsureDir(outDir)
    file := outDir "\timetrak_" label "_" FormatTime(, "HHmmss") ".png"
    ps1 := A_ScriptDir "\capture_screen.ps1"
    try {
        RunWait Format('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{1}" -OutFile "{2}"', ps1, file), , "Hide"
    } catch {
    }
}

Fail(msg) {
    global UNATTENDED
    SnapAutomationScreen("fail")
    if !UNATTENDED
        MsgBox msg
    ExitApp 1
}
