#!/usr/bin/env python3
"""
Automate eSSL eTimeTrackLite (native WinForms) with pywinauto.

Office flow:
  1) Login
  2) Utilities -> Device Management
     - select LGF/UGF devices (skip USB and UGF IN 1)
     - Start Download
     - wait until all selected devices finish
  3) Close all dialogs until the main window is plain
  4) Attendance Reports -> Daily Attendance Reports -> Detailed Attendance Report
  5) Export -> Save As: {Month} {Location}.xls (e.g. Aug Agra.xls from LOCATION_CODE)
  6) Close open pages, then Log Off (3rd toolbar icon)
  7) Click Close on the Login dialog to exit the app

Use regular `python.exe` (not free-threaded `py -3` / python3.13t).

Examples:
  python scripts/essl_export.py --login-only
  python scripts/essl_export.py --select-only
  python scripts/essl_export.py --sync-only
  python scripts/essl_export.py --skip-sync
  python scripts/essl_export.py --report monthly-basic
  python scripts/essl_export.py --report detailed
  python scripts/essl_export.py
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
import shutil

try:
    from pywinauto import Application, Desktop, mouse
    from pywinauto.keyboard import send_keys
except ImportError:
    sys.exit("Missing dependency: pywinauto. Install with: python -m pip install pywinauto")

REPO_ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = REPO_ROOT / "logs"
UI_DUMP_DIR = REPO_ROOT / "output" / "essl_ui"

APP_EXE_CANDIDATES = [
    Path(r"C:\Program Files (x86)\essl\eTimeTrackLite\eTimeTrackLite.exe"),
    Path(r"C:\Program Files (x86)\eSSL\eTimeTrackLite\eTimeTrackLite.exe"),
    Path(r"C:\Program Files\essl\eTimeTrackLite\eTimeTrackLite.exe"),
]

DEFAULT_USER = os.getenv("ESSL_USER", "essl")
DEFAULT_PASSWORD = os.getenv("ESSL_PASSWORD", "essl")
DEFAULT_EXPORT_DIR = Path(os.getenv("ATTENDANCE_DIR", r"D:\Attendance"))
DEFAULT_DEVICES_CSV = "LGF OUT,LGF IN,UGF OUT"
DEFAULT_SKIP_CSV = "USB,UGF IN 1"
# Office request: Daily -> Detailed Attendance Report
DEFAULT_REPORT = os.getenv("ESSL_REPORT", "monthly-basic").strip().lower()

LOCATION_EXPORT_NAMES = {
    "AGRA": "Agra",
    "NOIDA": "Noida",
    "HYD": "Hyd",
    "HYDERABAD": "Hyd",
}

LOGGER = logging.getLogger("essl_export")

TOP_LEVEL_MENU = {
    "Admin",
    "Masters",
    "Utilities",
    "Attendance Reports",
    "Canteen",
    "Monitor",
    "Windows",
    "Help",
    "Close",
    "Restore",
    "Minimize",
    "System",
    "Move",
    "Size",
    "Maximize",
}

DONE_STATUS_RE = re.compile(
    r"complete|completed|done|success|downloaded|finished|ok",
    re.I,
)
BUSY_STATUS_RE = re.compile(
    r"download|connecting|progress|running|busy|please wait|in progress",
    re.I,
)


def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_DIR / "essl_export.log", encoding="utf-8"),
        ],
    )


def load_dotenv() -> None:
    env_path = REPO_ROOT / ".env"
    if not env_path.exists():
        return
    text = env_path.read_text(encoding="utf-8-sig", errors="ignore")
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def find_exe() -> Path:
    for path in APP_EXE_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "eTimeTrackLite.exe not found. Update APP_EXE_CANDIDATES in essl_export.py."
    )


def as_window(wrapper):
    handle = wrapper.handle
    app = Application(backend="uia").connect(handle=handle)
    return app.window(handle=handle)


def find_main_window(timeout: float = 45.0):
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            for w in Desktop(backend="uia").windows():
                title = w.window_text() or ""
                if "eTimeTrackLite" in title and "Login" not in title and "GDI+" not in title:
                    return as_window(w)
            # Fallback: win32 top-level title match
            for w in Desktop(backend="win32").windows():
                title = w.window_text() or ""
                if "eTimeTrackLite" in title and "Login" not in title and "GDI+" not in title:
                    return as_window(w)
        except Exception as exc:
            last_err = exc
        time.sleep(0.5)
    msg = "eTimeTrackLite main window did not appear."
    if last_err:
        msg += f" Last error: {last_err}"
    raise TimeoutError(msg)


def find_login_window(timeout: float = 8.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for w in Desktop(backend="uia").windows():
            title = w.window_text() or ""
            if "Login" in title and "eTimeTrack" in title:
                return as_window(w)
        try:
            for w in Desktop(backend="uia").windows():
                title = w.window_text() or ""
                if "eTimeTrackLite" not in title or "Login" in title:
                    continue
                for child in w.descendants(control_type="Window"):
                    if "Login" in (child.window_text() or ""):
                        return as_window(child)
                named = w.child_window(auto_id="frmLogin", control_type="Window")
                if named.exists(timeout=0.2):
                    return as_window(named)
        except Exception:
            pass
        time.sleep(0.4)
    return None


def connect_app(exe: Path) -> Application:
    existing = None
    for w in Desktop(backend="uia").windows():
        title = w.window_text() or ""
        if "eTimeTrackLite" in title and "GDI+" not in title:
            existing = w
            break

    if existing is None:
        LOGGER.info("Starting %s", exe)
        app = Application(backend="uia").start(f'"{exe}"')
        time.sleep(3)
        return app

    LOGGER.info("Connecting to running eTimeTrackLite (%s)", existing.window_text())
    return Application(backend="uia").connect(handle=existing.handle)


def dump_ui(path: Path | None = None) -> Path:
    UI_DUMP_DIR.mkdir(parents=True, exist_ok=True)
    out = path or (UI_DUMP_DIR / f"tree_{datetime.now():%Y%m%d_%H%M%S}.txt")
    win = find_main_window(timeout=10)
    win.set_focus()
    win.print_control_identifiers(filename=str(out), depth=6)
    LOGGER.info("UI dump written to %s", out)
    return out


def do_login(user: str, password: str) -> None:
    login = find_login_window(timeout=12)
    if login is None:
        LOGGER.info("No login dialog - already logged in.")
        return

    LOGGER.info("Filling login as %s", user)
    login.set_focus()
    time.sleep(0.4)
    try:
        login.child_window(auto_id="txt_LoginName", control_type="Edit").set_edit_text(user)
        login.child_window(auto_id="txt_password", control_type="Edit").set_edit_text(password)
        login.child_window(auto_id="btn_Login", control_type="Button").click_input()
    except Exception:
        edits = login.descendants(control_type="Edit")
        if len(edits) < 2:
            raise RuntimeError("Login fields not found.")
        edits[0].set_edit_text(user)
        edits[1].set_edit_text(password)
        login.child_window(title="Login", control_type="Button").click_input()

    deadline = time.time() + 25
    while time.time() < deadline:
        if find_login_window(timeout=0.8) is None:
            LOGGER.info("Login succeeded.")
            time.sleep(1.0)
            return
        time.sleep(0.5)
    raise TimeoutError("Login dialog stayed open - check credentials.")


def visible_submenu_items(main):
    items = []
    for c in main.descendants(control_type="MenuItem"):
        title = (c.window_text() or "").strip()
        try:
            rect = c.rectangle()
        except Exception:
            continue
        area = rect.width() * rect.height()
        if title and area > 800 and rect.top >= 40 and title not in TOP_LEVEL_MENU:
            items.append((title, rect, c))
    return items


def close_mdi_children(main, max_rounds: int = 8) -> None:
    """Close open MDI / report pages (Ctrl+F4)."""
    main.set_focus()
    for _ in range(max_rounds):
        title = main.window_text() or ""
        if "[" not in title:
            break
        send_keys("^{F4}")
        time.sleep(0.4)
    send_keys("{ESC}{ESC}")
    time.sleep(0.2)


def _click_dialog_close_buttons(main) -> bool:
    """Click Close only on filter dialogs that also have Generate (e.g. Monthly/Daily Status Report)."""
    for c in list(main.descendants(control_type="Window")):
        title = (c.window_text() or "").strip()
        if not title or "eTimeTrackLite" in title:
            continue
        buttons = list(c.descendants(control_type="Button"))
        labels = [(b.window_text() or "").strip() for b in buttons]
        if "Generate" not in labels or "Close" not in labels:
            continue
        for b in buttons:
            if (b.window_text() or "").strip() != "Close":
                continue
            try:
                LOGGER.info("Closing filter dialog: %s", title)
                b.click_input()
                time.sleep(0.6)
                return True
            except Exception:
                continue
    return False


def ensure_plain_ui(main, max_rounds: int = 12):
    """
    Close every open dialog/MDI child until the main window is plain
    (title without [ ... ], no Report filter dialogs).
    """
    LOGGER.info("Closing dialogs until UI is plain")
    for round_i in range(1, max_rounds + 1):
        main = _connect_main_fast() or find_main_window(timeout=10)
        try:
            main.set_focus()
        except Exception:
            pass
        send_keys("{ESC}")
        time.sleep(0.1)

        title = main.window_text() or ""

        # 1) Close filter dialogs (Monthly Status Report / Daily Attendance Report filters)
        if _click_dialog_close_buttons(main):
            continue

        # 2) Close MDI report viewers / Device Management tabs
        if "[" in title:
            LOGGER.info("Closing MDI child via Ctrl+F4: %s", title)
            send_keys("^{F4}")
            time.sleep(0.5)
            continue

        # 3) Any leftover named child windows?
        leftover_title = None
        for c in main.descendants(control_type="Window"):
            ct = (c.window_text() or "").strip()
            if ct and any(k in ct for k in ("Report", "Device Management", "Users List", "Employee")):
                leftover_title = ct
                break
        if leftover_title:
            LOGGER.info("Closing leftover child via Ctrl+F4: %s", leftover_title)
            send_keys("^{F4}")
            time.sleep(0.5)
            continue

        title = main.window_text() or ""
        if "[" not in title:
            LOGGER.info("UI is plain: %s", title)
            return main

        LOGGER.info("Plain check round %s still: %s", round_i, title)
        send_keys("^{F4}")
        time.sleep(0.4)

    main = _connect_main_fast() or find_main_window(timeout=10)
    LOGGER.warning("UI may not be fully plain: %s", main.window_text())
    return main


def close_all_pages(main) -> None:
    """Close report viewers and dialogs (used before logout)."""
    return ensure_plain_ui(main)


def do_logout(main) -> None:
    """After UI is plain, click the 3rd toolbar icon (Log Off / door+arrow)."""
    LOGGER.info("Logging off via 3rd toolbar icon (Log Off)")
    main = _connect_main_fast() or main
    try:
        main.set_focus()
    except Exception:
        pass
    time.sleep(0.4)

    btn = None
    # Prefer exact 3rd button on ToolStrip1 (Users, Change Password, Log Off)
    try:
        toolbar = main.child_window(auto_id="ToolStrip1", control_type="ToolBar")
        buttons = []
        for b in toolbar.descendants(control_type="Button"):
            try:
                r = b.rectangle()
            except Exception:
                continue
            # Top toolbar row only
            if r.top < 90 and r.height() <= 40:
                buttons.append((r.left, b))
        buttons.sort(key=lambda x: x[0])
        if len(buttons) >= 3:
            btn = buttons[2][1]
            LOGGER.info(
                "Using 3rd toolbar button: %s",
                (btn.window_text() or "").strip() or "(icon)",
            )
    except Exception as exc:
        LOGGER.warning("Toolbar scan failed: %s", exc)

    if btn is None:
        btn = first_by_title(main, "Log Off", "Button")

    if btn is None:
        raise RuntimeError("Log Off (3rd toolbar icon) not found after closing dialogs.")

    btn.click_input()
    time.sleep(1.2)

    # Confirm Yes/OK if a prompt appears
    for w in Desktop(backend="uia").windows():
        title = (w.window_text() or "").lower()
        if any(k in title for k in ("log off", "logout", "confirm", "question", "attention")):
            for label in ("Yes", "OK", "Log Off"):
                try:
                    b = w.child_window(title=label, control_type="Button")
                    if b.exists(timeout=0.3):
                        b.click_input()
                        break
                except Exception:
                    pass
            else:
                send_keys("{ENTER}")
            time.sleep(0.5)
            break

    LOGGER.info("Log Off clicked")


def close_login_or_app() -> None:
    """
    After Log Off, the Login dialog is shown. Click Close on it to exit the app.
    Falls back to closing the main process window.
    """
    LOGGER.info("Clicking Close on login dialog (final step)")
    time.sleep(0.8)

    login = find_login_window(timeout=8)
    if login is not None:
        try:
            login.set_focus()
        except Exception:
            pass
        time.sleep(0.3)
        try:
            close_btn = login.child_window(auto_id="btn_Close", control_type="Button")
            if close_btn.exists(timeout=1):
                close_btn.click_input()
                time.sleep(1.0)
                LOGGER.info("Login dialog Close clicked")
                return
        except Exception:
            pass
        btn = first_by_title(login, "Close", "Button")
        if btn is not None:
            btn.click_input()
            time.sleep(1.0)
            LOGGER.info("Login dialog Close clicked")
            return

    # Fallback: Alt+F4 on any remaining eTimeTrack window
    main = _connect_main_fast()
    if main is not None:
        try:
            main.set_focus()
            send_keys("%{F4}")
            time.sleep(1.0)
            LOGGER.info("Sent Alt+F4 to close application")
            return
        except Exception:
            pass

    LOGGER.warning("Could not find Login Close button — app may still be open")


def menu_click(main, *path: str) -> None:
    main.set_focus()
    time.sleep(0.2)
    # Do not Ctrl+F4 here aggressively; caller should ensure_plain_ui first.
    send_keys("{ESC}{ESC}")
    time.sleep(0.2)

    if path[:3] == (
        "Attendance Reports",
        "Monthly Reports",
        "Monthly Basic Report",
    ):
        LOGGER.info("Opening Monthly Basic Report via Alt menu navigation")
        send_keys("%")
        time.sleep(0.3)
        send_keys("{RIGHT}{RIGHT}{RIGHT}")  # Attendance Reports
        time.sleep(0.3)
        send_keys("{DOWN}")
        time.sleep(0.35)
        send_keys("{DOWN}")  # Daily
        time.sleep(0.2)
        send_keys("{DOWN}")  # Monthly Reports
        time.sleep(0.25)
        send_keys("{RIGHT}")
        time.sleep(0.35)
        send_keys("{ENTER}")  # Monthly Basic Report
        time.sleep(1.2)
        return

    if path[:3] == (
        "Attendance Reports",
        "Daily Attendance Reports",
        "Basic Attendance Report",
    ):
        LOGGER.info("Opening Daily Basic Attendance Report via Alt menu navigation")
        send_keys("%")
        time.sleep(0.3)
        send_keys("{RIGHT}{RIGHT}{RIGHT}")
        time.sleep(0.3)
        send_keys("{DOWN}")
        time.sleep(0.35)
        send_keys("{DOWN}")
        time.sleep(0.25)
        send_keys("{RIGHT}")
        time.sleep(0.35)
        send_keys("{ENTER}")
        time.sleep(1.2)
        return

    if path[:3] == (
        "Attendance Reports",
        "Daily Attendance Reports",
        "Detailed Attendance Report",
    ):
        LOGGER.info("Opening Daily Detailed Attendance Report via Alt menu navigation")
        send_keys("%")
        time.sleep(0.3)
        send_keys("{RIGHT}{RIGHT}{RIGHT}")
        time.sleep(0.3)
        send_keys("{DOWN}")
        time.sleep(0.35)
        send_keys("{DOWN}")
        time.sleep(0.25)
        send_keys("{RIGHT}")
        time.sleep(0.3)
        send_keys("{DOWN}")  # Basic -> Detailed
        time.sleep(0.25)
        send_keys("{ENTER}")
        time.sleep(1.2)
        return

    if path[:2] == ("Utilities", "Device Management"):
        LOGGER.info("Opening Device Management via Alt menu navigation")
        send_keys("%")
        time.sleep(0.3)
        send_keys("{RIGHT}{RIGHT}")
        time.sleep(0.3)
        send_keys("{DOWN}")
        time.sleep(0.25)
        send_keys("{ENTER}")
        time.sleep(1.0)
        return

    menu = main.child_window(auto_id="MenuStrip1", control_type="MenuBar")
    top = menu.child_window(title=path[0], control_type="MenuItem")
    top.click_input()
    time.sleep(0.45)
    for name in path[1:]:
        matches = [c for t, _, c in visible_submenu_items(main) if t == name]
        if not matches:
            raise RuntimeError(f"Menu item not visible: {name}")
        matches[0].click_input()
        time.sleep(0.45)
    time.sleep(0.6)


def open_device_management(main) -> None:
    LOGGER.info("Opening Utilities -> Device Management")
    for attempt in range(1, 4):
        menu_click(main, "Utilities", "Device Management")
        time.sleep(1.0)
        main = find_main_window(timeout=10)
        title = main.window_text() or ""
        if "Device Management" in title:
            LOGGER.info("Device Management open: %s", title)
            return
        # Also accept if Start Download exists even when title lags
        if first_by_title(main, "Start Download", "Button") is not None:
            LOGGER.info("Device Management controls visible (attempt %s)", attempt)
            return
        LOGGER.warning("Device Management not open yet (attempt %s): %s", attempt, title)
        time.sleep(0.5)
    raise TimeoutError("Could not open Device Management")


def click_start_download(main) -> None:
    LOGGER.info("Clicking Start Download")
    main = _connect_main_fast() or main
    try:
        main.set_focus()
    except Exception:
        pass
    time.sleep(0.3)

    btn = first_by_title(main, "Start Download", "Button")
    if btn is None:
        time.sleep(1.0)
        main = _connect_main_fast() or find_main_window(timeout=10)
        btn = first_by_title(main, "Start Download", "Button")

    if btn is None:
        LOGGER.warning("Start Download not found by title — trying coordinate click")
        r = main.rectangle()
        mouse.click(coords=(r.right - 120, r.top + 110))
        time.sleep(1.0)
        LOGGER.info("Start Download coordinate click sent")
        return

    btn.click_input()
    time.sleep(1.0)
    LOGGER.info("Start Download clicked - waiting for device sync to finish")


def _device_name(ctrl) -> str:
    try:
        return (ctrl.window_text() or "").strip()
    except Exception:
        return ""


def first_by_title(parent, title: str, control_type: str):
    matches = [
        c
        for c in parent.descendants(control_type=control_type)
        if (c.window_text() or "") == title
    ]
    if not matches:
        return None
    matches.sort(key=lambda c: (c.rectangle().width() * c.rectangle().height()), reverse=True)
    return matches[0]


def _row_device_name(main, idx: int) -> str:
    """Read Device Name cell value for row idx (UIA name is only a label)."""
    cells = [
        c
        for c in main.descendants(control_type="DataItem")
        if (c.window_text() or "") == f"Device Name Row {idx}"
    ]
    if not cells:
        known = ["USB", "LGF OUT", "LGF IN", "UGF OUT", "UGF IN 1"]
        return known[idx] if 0 <= idx < len(known) else ""
    cell = cells[0]
    try:
        if hasattr(cell, "iface_value") and cell.iface_value is not None:
            value = str(cell.iface_value.CurrentValue or "").strip()
            if value:
                return value
    except Exception:
        pass
    try:
        legacy = cell.legacy_properties() or {}
        value = str(legacy.get("Value") or "").strip()
        if value:
            return value
    except Exception:
        pass
    known = ["USB", "LGF OUT", "LGF IN", "UGF OUT", "UGF IN 1"]
    return known[idx] if 0 <= idx < len(known) else ""


def _click_row_checkbox(main, idx: int) -> bool:
    """Click the checkbox cell for Device List row idx."""
    cells = [
        c
        for c in main.descendants(control_type="DataItem")
        if (c.window_text() or "") == f" Row {idx}"
    ]
    if not cells:
        return False
    cells.sort(key=lambda c: c.rectangle().left)
    cell = cells[0]
    try:
        r = cell.rectangle()
        mouse.click(coords=(r.left + max(6, r.width() // 2), r.top + max(6, r.height() // 2)))
        return True
    except Exception:
        try:
            cell.click_input()
            return True
        except Exception:
            return False


def select_devices(main, wanted: list[str], skip: list[str]) -> None:
    """
    Select-all via header checkbox, then uncheck skipped devices (USB, UGF IN 1, ...).
    Confirmed office layout: USB, LGF OUT, LGF IN, UGF OUT, UGF IN 1.
    """
    skip_upper = {s.strip().upper() for s in skip if s.strip()}
    LOGGER.info("Selecting devices (wanted=%s skip=%s)", wanted, sorted(skip_upper))
    main = _connect_main_fast() or main
    main.set_focus()
    time.sleep(0.5)

    header = None
    for c in main.descendants(control_type="CheckBox"):
        name = _device_name(c)
        try:
            r = c.rectangle()
        except Exception:
            continue
        # Header select-all is the tiny unnamed checkbox above the grid
        if not name and r.top < 140 and r.left < 40 and r.width() <= 24:
            header = c
            break

    if header is not None:
        LOGGER.info("Clicking header select-all checkbox")
        header.click_input()
        time.sleep(0.8)
    else:
        LOGGER.warning("Header checkbox not found")

    # Reconnect after select-all — UIA tree goes stale and row DataItems disappear otherwise
    main = _connect_main_fast() or main
    time.sleep(0.4)

    unchecked = []
    for idx in range(0, 12):
        main = _connect_main_fast() or main
        cells = [
            c
            for c in main.descendants(control_type="DataItem")
            if (c.window_text() or "") == f" Row {idx}"
        ]
        if not cells:
            if idx == 0:
                LOGGER.warning("No device rows found after select-all")
            break
        device = _row_device_name(main, idx)
        LOGGER.info("Row %s device=%r", idx, device)
        if device.upper() in skip_upper:
            if _click_row_checkbox(main, idx):
                LOGGER.info("Unchecked row %s (%s)", idx, device)
                unchecked.append(device)
                time.sleep(0.45)
            else:
                LOGGER.warning("Failed to uncheck row %s (%s)", idx, device)

    if not unchecked and skip_upper:
        LOGGER.warning("No skip devices were unchecked (expected %s)", sorted(skip_upper))
    else:
        LOGGER.info("Skip devices unchecked: %s", unchecked)

    if first_by_title(main, "Start Download", "Button") is not None:
        LOGGER.info("Device Management ready (Start Download visible)")
    else:
        LOGGER.warning("Start Download button not visible yet")


def _status_texts(main) -> list[str]:
    texts = []
    for idx in range(0, 12):
        cells = [
            c
            for c in main.descendants(control_type="DataItem")
            if (c.window_text() or "") == f"Status Row {idx}"
        ]
        if not cells:
            # Some builds expose name as the value itself
            break
        cell = cells[0]
        value = ""
        try:
            value = (cell.window_text() or "").strip()
        except Exception:
            pass
        if value in ("", f"Status Row {idx}"):
            try:
                value = (getattr(cell.element_info, "name", None) or "").strip()
            except Exception:
                value = ""
        if value in ("", f"Status Row {idx}"):
            try:
                # Legacy / value pattern fallbacks
                value = str(cell.get_value() if hasattr(cell, "get_value") else "")
            except Exception:
                value = ""
        texts.append(value)
    return texts


def _download_busy(main) -> bool:
    stop = first_by_title(main, "Stop Download", "Button")
    if stop is not None:
        try:
            if stop.is_enabled():
                return True
        except Exception:
            pass
    start = first_by_title(main, "Start Download", "Button")
    if start is not None:
        try:
            if not start.is_enabled():
                return True
        except Exception:
            pass
    return False


def _connect_main_fast():
    """Reconnect to eTimeTrackLite without enumerating every desktop window."""
    # Fastest: attach by process executable name
    try:
        app = Application(backend="uia").connect(path="eTimeTrackLite.exe", timeout=2)
        wins = app.windows()
        for w in wins:
            title = w.window_text() or ""
            if "eTimeTrackLite" in title and "Login" not in title:
                return as_window(w)
        if wins:
            return as_window(wins[0])
    except Exception:
        pass
    # win32 by title substring (lighter than full UIA desktop scan)
    try:
        app = Application(backend="win32").connect(title_re=".*eTimeTrackLite.*", timeout=2)
        w = app.top_window()
        return as_window(w)
    except Exception:
        pass
    return None


def wait_download(main, timeout_sec: int, poll_sec: int = 5) -> None:
    """Wait until device download finishes (not just a fixed sleep)."""
    LOGGER.info("Waiting for device download to finish (timeout %ss)...", timeout_sec)
    elapsed = 0
    saw_busy = False
    stable_done = 0

    while elapsed < timeout_sec:
        time.sleep(poll_sec)
        elapsed += poll_sec

        # During download, full Desktop UIA scans can hang — use process connect.
        refreshed = _connect_main_fast()
        if refreshed is not None:
            main = refreshed
        else:
            LOGGER.info("...%ss (reconnect deferred — still waiting)", elapsed)
            continue

        try:
            main.set_focus()
        except Exception:
            pass

        busy = _download_busy(main)
        statuses = _status_texts(main)
        status_blob = " | ".join(s for s in statuses if s)
        if status_blob:
            LOGGER.info("...%ss status=[%s] busy=%s", elapsed, status_blob, busy)
        else:
            LOGGER.info("...%ss busy=%s", elapsed, busy)

        if busy:
            saw_busy = True
            stable_done = 0
            continue

        useful = [s for s in statuses[1:] if s and not s.lower().startswith("status row")]
        if useful:
            if any(BUSY_STATUS_RE.search(s) for s in useful):
                stable_done = 0
                continue
            if all(DONE_STATUS_RE.search(s) for s in useful) or all(s.strip() for s in useful):
                stable_done += 1
            else:
                stable_done = 0
        elif saw_busy and not busy:
            stable_done += 1
        elif elapsed >= min(45, timeout_sec) and not busy:
            # Devices may finish with empty Status cells; require idle for a while
            stable_done += 1
        else:
            stable_done = 0

        if stable_done >= 2:
            LOGGER.info("Device download finished after %ss", elapsed)
            return

    raise TimeoutError(
        f"Device download did not finish within {timeout_sec}s. "
        "Check Status / Logs Downloaded in Device Management."
    )


def open_report(main, report: str) -> None:
    report = (report or DEFAULT_REPORT).lower()
    if report in ("detailed", "daily-detailed", "daily_detailed"):
        LOGGER.info(
            "Opening Attendance Reports -> Daily Attendance Reports -> Detailed Attendance Report"
        )
        menu_click(
            main,
            "Attendance Reports",
            "Daily Attendance Reports",
            "Detailed Attendance Report",
        )
    elif report in ("monthly-basic", "monthly", "monthly_basic"):
        LOGGER.info(
            "Opening Attendance Reports -> Monthly Reports -> Monthly Basic/Status Report"
        )
        menu_click(
            main,
            "Attendance Reports",
            "Monthly Reports",
            "Monthly Basic Report",
        )
    else:
        LOGGER.info(
            "Opening Attendance Reports -> Daily Attendance Reports -> Basic Attendance Report"
        )
        menu_click(
            main,
            "Attendance Reports",
            "Daily Attendance Reports",
            "Basic Attendance Report",
        )
    time.sleep(1.5)


def _select_combo_value(win, wanted: str) -> bool:
    """Select an item in a WinForms ComboBox (Report Type, etc.)."""
    wanted_norm = re.sub(r"\s+", " ", wanted.strip().lower())
    combos = list(win.descendants(control_type="ComboBox"))
    for cb in combos:
        try:
            current = (cb.window_text() or "").strip()
        except Exception:
            current = ""
        if re.sub(r"\s+", " ", current.lower()) == wanted_norm:
            LOGGER.info("Report Type already set: %s", current)
            return True
        try:
            cb.select(wanted)
            time.sleep(0.4)
            LOGGER.info("Selected ComboBox value: %s", wanted)
            return True
        except Exception:
            pass
        try:
            cb.click_input()
            time.sleep(0.35)
            # Prefer exact list item click when exposed
            items = [
                i
                for i in win.descendants(control_type="ListItem")
                if re.sub(r"\s+", " ", (i.window_text() or "").strip().lower()) == wanted_norm
            ]
            if items:
                items[0].click_input()
                time.sleep(0.35)
                LOGGER.info("Clicked list item: %s", wanted)
                return True
            send_keys("^a")
            time.sleep(0.1)
            send_keys(wanted, with_spaces=True)
            time.sleep(0.2)
            send_keys("{ENTER}")
            time.sleep(0.4)
            LOGGER.info("Typed ComboBox value: %s", wanted)
            return True
        except Exception as exc:
            LOGGER.warning("ComboBox select failed: %s", exc)
    return False


def select_monthly_basic_work_duration(win) -> None:
    """On Monthly Status Report filter, set Report Type = Basic Work Duration (In/Out grid)."""
    title = (win.window_text() or "")
    LOGGER.info("Configuring monthly report filter: %s", title)
    # Prefer Basic Work Duration (Status + InTime + OutTime + Total) over plain Basic Report.
    report_type_ok = (
        _select_combo_value(win, "Basic Work Duration")
        or _select_combo_value(win, "Basic Report")
    )
    if not report_type_ok:
        main = find_main_window(timeout=5)
        if main is not None:
            report_type_ok = (
                _select_combo_value(main, "Basic Work Duration")
                or _select_combo_value(main, "Basic Report")
            )
        if report_type_ok:
            LOGGER.info("Set Report Type via main window")
        else:
            LOGGER.warning(
                "Could not set Report Type to Basic Work Duration — Generate may use current type"
            )
    for c in win.descendants(control_type="CheckBox"):
        name = (c.window_text() or "").strip().lower()
        if "recalculate" in name:
            try:
                LOGGER.info("Recalculate checkbox present: %s", c.window_text())
            except Exception:
                pass
            break


def _format_essl_date(day) -> str:
    """eSSL date pickers show values like '18 Sep 2026' / '01 Sep 2026'."""
    return day.strftime("%d %b %Y")


def report_export_date(from_date=None):
    """
    Calendar day used for From Date = To Date on the eSSL filter.

    Agreed office policy: use **today** (local PC date) for both the 1 AM and
    1 PM runs so the API gets the current day's punches (1 PM upserts over 1 AM).

    Overrides (edge cases / testing):
      - Explicit ``from_date`` argument
      - Env ``ESSL_REPORT_DATE=YYYY-MM-DD`` (single day, From=To)
    """
    override = (os.getenv("ESSL_REPORT_DATE") or "").strip()
    if override:
        try:
            return datetime.strptime(override, "%Y-%m-%d").date()
        except ValueError:
            LOGGER.warning(
                "Invalid ESSL_REPORT_DATE=%r (expected YYYY-MM-DD); using today",
                override,
            )
    return from_date or datetime.now().date()


def previous_weekday(from_date=None):
    """Deprecated: kept for manual/debug calls. Prefer report_export_date()."""
    day = (from_date or datetime.now().date()) - timedelta(days=1)
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def set_report_date_range(win, from_day=None, to_day=None) -> None:
    """
    Set From Date = To Date = today (local date), unless overridden.

    1 AM and 1 PM schedules both export the same calendar day; the API upserts
    so the afternoon run refreshes incomplete morning punches.
    """
    target = report_export_date()
    from_day = from_day or target
    to_day = to_day or from_day
    if to_day < from_day:
        LOGGER.warning(
            "To Date %s is before From Date %s — swapping to keep a valid range",
            to_day,
            from_day,
        )
        from_day, to_day = to_day, from_day
    from_text = _format_essl_date(from_day)
    to_text = _format_essl_date(to_day)
    LOGGER.info(
        "Setting report dates: From=%s To=%s (today / report day)",
        from_text,
        to_text,
    )

    # Collect editable date fields (DateTimePicker / Edit / ComboBox near date labels)
    candidates = []
    for ctype in ("Edit", "ComboBox", "Spinner", "Pane"):
        for c in win.descendants(control_type=ctype):
            try:
                name = (c.window_text() or "").strip()
                auto = ""
                try:
                    auto = (c.element_info.automation_id or "") + " " + (c.element_info.name or "")
                except Exception:
                    pass
                hay = f"{name} {auto}".lower()
                # Likely a date value already (contains month abbr) or empty editable near dates
                if any(m in hay for m in (
                    "jan", "feb", "mar", "apr", "may", "jun",
                    "jul", "aug", "sep", "oct", "nov", "dec",
                )) or (ctype == "Edit" and name):
                    r = c.rectangle()
                    candidates.append((r.top, r.left, c, name, ctype))
            except Exception:
                continue

    candidates.sort()  # top-to-bottom, then left-to-right
    # Prefer the two topmost date-looking fields (From then To on the same row)
    date_fields = []
    for top, left, ctrl, name, ctype in candidates:
        if any(m in name.lower() for m in (
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec",
        )):
            date_fields.append((top, left, ctrl, name, ctype))
    date_fields.sort(key=lambda x: (x[0], x[1]))

    if len(date_fields) < 2:
        # Fallback: click labels then type
        LOGGER.warning("Could not find two date fields by value — trying label-relative Edits")
        date_fields = []
        for c in win.descendants(control_type="Text"):
            label = (c.window_text() or "").strip().lower()
            if label not in {"from date", "to date", "from", "to"}:
                continue
            try:
                lr = c.rectangle()
            except Exception:
                continue
            # Nearest Edit/Combo to the right of the label
            best = None
            best_dist = 10_000
            for ctype in ("Edit", "ComboBox"):
                for e in win.descendants(control_type=ctype):
                    try:
                        er = e.rectangle()
                    except Exception:
                        continue
                    if er.left < lr.left:
                        continue
                    if abs(er.top - lr.top) > 20:
                        continue
                    dist = er.left - lr.left
                    if dist < best_dist:
                        best_dist = dist
                        best = (er.top, er.left, e, e.window_text() or "", ctype)
            if best:
                date_fields.append(best)
        date_fields.sort(key=lambda x: (0 if "from" in str(x) else 1, x[0], x[1]))

    if len(date_fields) < 2:
        LOGGER.warning("Date fields not found — leaving From/To as shown in dialog")
        return

    # First = From (left), second = To (right) on the top row
    top_row = sorted(date_fields[:4], key=lambda x: (x[0], x[1]))
    from_ctrl = top_row[0][2]
    to_ctrl = top_row[1][2] if len(top_row) > 1 else top_row[0][2]

    def _set_date_ctrl(ctrl, value: str, label: str) -> bool:
        try:
            ctrl.set_focus()
            time.sleep(0.15)
            # Many DateTimePickers expose an Edit child
            try:
                ctrl.set_edit_text(value)
                LOGGER.info("Set %s via set_edit_text -> %s", label, value)
                return True
            except Exception:
                pass
            send_keys("^a")
            time.sleep(0.05)
            send_keys(value, with_spaces=True)
            send_keys("{ENTER}")
            time.sleep(0.2)
            LOGGER.info("Set %s via keyboard -> %s", label, value)
            return True
        except Exception as exc:
            LOGGER.warning("Failed to set %s to %s: %s", label, value, exc)
            return False

    _set_date_ctrl(from_ctrl, from_text, "From Date")
    time.sleep(0.25)
    _set_date_ctrl(to_ctrl, to_text, "To Date")
    time.sleep(0.25)


def dismiss_info_dialogs(timeout: float = 5.0) -> None:
    """Close eSSL Info / warning popups (e.g. Please select Atleast One Company)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        closed = False
        try:
            for w in Desktop(backend="uia").windows():
                title = (w.window_text() or "").strip()
                # Only touch small eSSL message boxes — never Cursor/IDE windows.
                if title not in {"Info", "Warning", "Error", "Message"}:
                    continue
                try:
                    w.set_focus()
                    for b in w.descendants(control_type="Button"):
                        if (b.window_text() or "").strip() in {"OK", "Ok", "Close"}:
                            b.click_input()
                            LOGGER.info("Dismissed dialog: %s", title or "Info")
                            closed = True
                            time.sleep(0.4)
                            break
                except Exception:
                    pass
        except Exception:
            pass
        if not closed:
            break


def select_walking_tree_company(win) -> None:
    """
    On Monthly Status Report filter:
      Filter Company → Deselect All → select WalkingTree
    so the Excel only contains Walking Tree employees (smaller payload).
    """
    company_name = os.getenv("ESSL_COMPANY", "WalkingTree").strip() or "WalkingTree"
    # Accept "Walking Tree" or "WalkingTree"
    company_key = company_name.replace(" ", "").lower()
    LOGGER.info("Filtering company to: %s", company_name)

    # 1) Enable Filter Company
    for c in win.descendants(control_type="CheckBox"):
        name = (c.window_text() or "").strip().lower()
        if "filter company" in name:
            try:
                if c.get_toggle_state() == 0:
                    c.click_input()
                    time.sleep(0.4)
                    LOGGER.info("Enabled Filter Company")
            except Exception as exc:
                LOGGER.warning("Could not toggle Filter Company: %s", exc)
            break

    # 2) Deselect All on the COMPANY list only (left list, not Department)
    company_list = None
    for lst in win.descendants(control_type="List"):
        try:
            items = [i.window_text() or "" for i in lst.descendants(control_type="ListItem")]
        except Exception:
            continue
        joined = " ".join(items).lower().replace(" ", "")
        if "walkingtree" in joined or "contractor" in joined:
            company_list = lst
            break

    deselected = False
    if company_list is not None:
        try:
            list_rect = company_list.rectangle()
        except Exception:
            list_rect = None
        for c in win.descendants(control_type="RadioButton"):
            name = (c.window_text() or "").strip().lower()
            if name != "deselect all":
                continue
            try:
                r = c.rectangle()
                # Company Deselect All sits under the company list (x near list left)
                if list_rect is not None and abs(r.left - list_rect.left) > 80:
                    continue
                c.click_input()
                time.sleep(0.4)
                LOGGER.info("Clicked company Deselect All")
                deselected = True
                break
            except Exception as exc:
                LOGGER.warning("Deselect All click failed: %s", exc)
    if not deselected:
        # Fallback: first enabled Deselect All
        for c in win.descendants(control_type="RadioButton"):
            if (c.window_text() or "").strip().lower() != "deselect all":
                continue
            try:
                c.click_input()
                time.sleep(0.4)
                LOGGER.info("Clicked Deselect All (fallback)")
                deselected = True
                break
            except Exception:
                pass
    if not deselected:
        LOGGER.warning("Deselect All radio not found")

    # 3) Select WalkingTree in the company list
    selected = False
    search_roots = [company_list] if company_list is not None else [win]
    for root in search_roots:
        if root is None:
            continue
        for c in root.descendants(control_type="ListItem"):
            name = (c.window_text() or "").strip()
            if name.replace(" ", "").lower() != company_key:
                continue
            try:
                c.click_input()
                time.sleep(0.4)
                LOGGER.info("Selected company: %s", name)
                selected = True
                break
            except Exception as exc:
                LOGGER.warning("Could not click company %s: %s", name, exc)
        if selected:
            break

    if not selected:
        # Last resort: any ListItem matching across the dialog
        for c in win.descendants(control_type="ListItem"):
            name = (c.window_text() or "").strip()
            if name.replace(" ", "").lower() == company_key:
                try:
                    c.click_input()
                    time.sleep(0.4)
                    LOGGER.info("Selected company (global): %s", name)
                    selected = True
                    break
                except Exception as exc:
                    LOGGER.warning("Could not click company %s: %s", name, exc)

    if not selected:
        raise RuntimeError(
            f"Company '{company_name}' not found/selected in Filter Company list. "
            "Refusing to Generate (would show 'Please select Atleast One Company')."
        )


def generate_report(main, report: str | None = None, timeout: float = 30.0) -> None:
    report = (report or DEFAULT_REPORT).lower()
    LOGGER.info("Waiting for report filter / Generate")
    deadline = time.time() + timeout
    report_win = None
    while time.time() < deadline:
        for c in main.descendants(control_type="Window"):
            title = c.window_text() or ""
            if "Report" in title or "Attendance" in title or "Monthly" in title:
                # Prefer filter dialog that has Generate
                has_gen = any(
                    (b.window_text() or "") == "Generate"
                    for b in c.descendants(control_type="Button")
                )
                if has_gen:
                    report_win = as_window(c)
                    break
        if report_win:
            break
        time.sleep(0.5)
    if report_win is None:
        raise TimeoutError("Report filter window with Generate did not open.")

    win = report_win
    win.set_focus()
    time.sleep(0.5)

    if report in ("monthly-basic", "monthly", "monthly_basic"):
        select_monthly_basic_work_duration(win)
        time.sleep(0.3)
        set_report_date_range(win)  # From=To=today (1 AM & 1 PM runs)
        time.sleep(0.3)
        select_walking_tree_company(win)
        time.sleep(0.3)

    dismiss_info_dialogs(timeout=2.0)

    btn = first_by_title(win, "Generate", "Button") or first_by_title(main, "Generate", "Button")
    if btn is not None:
        LOGGER.info("Clicking Generate")
        btn.click_input()
    else:
        send_keys("%g")
        LOGGER.warning("Generate button not found - sent Alt+G.")

    time.sleep(0.8)
    dismiss_info_dialogs(timeout=3.0)

    deadline = time.time() + 120
    while time.time() < deadline:
        main_now = find_main_window(timeout=5)
        for c in main_now.descendants(control_type="MenuItem"):
            if (c.window_text() or "") == "Export":
                try:
                    r = c.rectangle()
                except Exception:
                    continue
                if r.top < 160 and r.width() * r.height() > 200:
                    LOGGER.info("Report viewer ready (Export visible)")
                    time.sleep(1.0)
                    return
        time.sleep(0.5)
    LOGGER.warning("Export toolbar not detected yet - continuing anyway")
    time.sleep(2)


PREVIOUS_FOLDER_MAX_FILES = 5


def _prune_previous_folder(archive_dir: Path, keep: int = PREVIOUS_FOLDER_MAX_FILES) -> None:
    """Keep only the newest `keep` files in previous\\; delete older ones."""
    if keep < 1 or not archive_dir.is_dir():
        return
    files = [p for p in archive_dir.iterdir() if p.is_file()]
    if len(files) <= keep:
        return
    # Newest first by mtime, then name for stable ties.
    files.sort(key=lambda p: (p.stat().st_mtime, p.name), reverse=True)
    for old in files[keep:]:
        try:
            old.unlink()
            LOGGER.info("Pruned old archive (keep %s): %s", keep, old.name)
        except Exception as exc:
            LOGGER.warning("Could not prune %s: %s", old, exc)


def archive_existing_export(target: Path) -> Path | None:
    """Move previous Excel aside so Save As does not block on Confirm Replace.

    Archived copies live in ATTENDANCE_DIR\\previous\\ and are capped at
    PREVIOUS_FOLDER_MAX_FILES (newest kept).
    """
    if not target.exists():
        return None
    archive_dir = target.parent / "previous"
    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = archive_dir / f"{target.stem}_{stamp}{target.suffix}"
    try:
        shutil.move(str(target), str(dest))
        LOGGER.info("Archived previous export -> %s", dest)
        _prune_previous_folder(archive_dir)
        return dest
    except Exception as exc:
        LOGGER.warning("Could not archive %s (%s) — will confirm replace on Save As", target, exc)
        try:
            target.unlink()
            LOGGER.info("Deleted previous export %s so new save can proceed", target)
        except Exception as exc2:
            LOGGER.warning("Could not delete previous export: %s", exc2)
        return None


def confirm_replace_if_prompted(timeout: float = 8.0) -> bool:
    """Click Yes on Windows 'Confirm Save As' / already-exists prompt."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for backend in ("uia", "win32"):
            try:
                windows = Desktop(backend=backend).windows()
            except Exception:
                continue
            for w in windows:
                title = (w.window_text() or "").strip()
                title_l = title.lower()
                if "confirm save as" not in title_l and title_l != "confirm":
                    # Also detect by body text
                    try:
                        body = " ".join((c.window_text() or "") for c in w.descendants())
                    except Exception:
                        body = ""
                    if "already exists" not in body.lower() and "do you want to replace" not in body.lower():
                        continue
                try:
                    w.set_focus()
                except Exception:
                    pass
                # Prefer clicking Yes button
                try:
                    for b in w.descendants(control_type="Button"):
                        name = (b.window_text() or "").strip().lower()
                        if name in {"yes", "&yes"}:
                            b.click_input()
                            LOGGER.info("Confirmed replace (clicked Yes)")
                            time.sleep(0.8)
                            return True
                except Exception:
                    pass
                send_keys("%y")
                LOGGER.info("Confirmed replace (Alt+Y)")
                time.sleep(0.8)
                return True
        time.sleep(0.3)
    return False


def export_excel(main, export_dir: Path, export_name: str) -> Path:
    export_dir.mkdir(parents=True, exist_ok=True)
    target = export_dir / export_name
    LOGGER.info("Exporting Excel to %s", target)

    # Move yesterday's / previous file out of the way before Save As.
    archive_existing_export(target)

    main.set_focus()
    time.sleep(0.3)

    # Focus generated report viewer (window without Generate)
    viewer = None
    best_area = 0
    for c in main.descendants(control_type="Window"):
        title = c.window_text() or ""
        if "Report" not in title and "Attendance" not in title:
            continue
        has_gen = any((b.window_text() or "") == "Generate" for b in c.descendants(control_type="Button"))
        if has_gen:
            continue
        try:
            r = c.rectangle()
        except Exception:
            continue
        area = r.width() * r.height()
        if area > best_area:
            best_area = area
            viewer = c
    if viewer is not None:
        try:
            viewer.set_focus()
            vr = viewer.rectangle()
            mouse.click(coords=(vr.left + 120, vr.top + 90))
            time.sleep(0.4)
        except Exception:
            pass

    export_item = None
    for c in main.descendants(control_type="MenuItem"):
        if (c.window_text() or "") == "Export":
            try:
                r = c.rectangle()
            except Exception:
                continue
            if r.top < 160 and r.width() * r.height() > 200:
                export_item = c
                break
    if export_item is None:
        raise RuntimeError("Report viewer Export control not found. Generate the report first.")

    LOGGER.info("Opening report Export (Save As)")
    rect = export_item.rectangle()
    mouse.click(coords=(rect.mid_point().x, rect.mid_point().y))
    time.sleep(0.8)

    popup = None
    deadline = time.time() + 5
    while time.time() < deadline and popup is None:
        for w in Desktop(backend="win32").windows():
            try:
                cls = w.class_name()
                wr = w.rectangle()
            except Exception:
                continue
            if "20808" in cls and wr.top >= 120 and wr.top <= 220 and wr.left >= 350:
                popup = w
                break
        time.sleep(0.2)

    if popup is not None:
        wr = popup.rectangle()
        LOGGER.info("Clicking Export flyout item at %s", wr)
        mouse.click(coords=(wr.left + 40, wr.top + 10))
        time.sleep(1.2)
    else:
        LOGGER.warning("Export flyout not found - trying right-edge click")
        mouse.click(coords=(rect.right - 1, rect.mid_point().y))
        time.sleep(1.2)

    save = None
    deadline = time.time() + 25
    while time.time() < deadline:
        for w in Desktop(backend="uia").windows():
            if "save" in (w.window_text() or "").lower():
                save = w
                break
        if save is None:
            for w in Desktop(backend="win32").windows():
                if (w.window_text() or "").startswith("Save"):
                    save = w
                    break
        if save:
            break
        time.sleep(0.4)
    if save is None:
        raise TimeoutError("Save As dialog did not appear after export.")

    try:
        save.set_focus()
    except Exception:
        pass
    time.sleep(0.4)

    # Correct Save As usage:
    # 1) Address / location bar = folder
    # 2) File name box = filename only
    LOGGER.info("Save As: folder=%s filename=%s", export_dir, export_name)
    filled = False
    try:
        dlg = Application(backend="win32").connect(handle=save.handle).window(handle=save.handle)
        dlg.set_focus()
        time.sleep(0.3)

        # Navigate folder via address bar
        send_keys("%d")
        time.sleep(0.4)
        send_keys("^a")
        time.sleep(0.1)
        send_keys(str(export_dir), with_spaces=True)
        send_keys("{ENTER}")
        time.sleep(1.2)

        # Set filename only in the File name edit
        edits = list(dlg.descendants(class_name="Edit"))
        chosen = None
        for edit in edits:
            text = (edit.window_text() or "")
            # Prefer the filename box (usually contains .xls / report), not toolbar path crumbs
            if ".xls" in text.lower() or "report" in text.lower() or text.endswith(".xlsx"):
                chosen = edit
                break
        if chosen is None:
            # Fall back to bottom-most edit (File name is usually lowest)
            ranked = []
            for edit in edits:
                try:
                    ranked.append((edit.rectangle().top, edit))
                except Exception:
                    pass
            ranked.sort(reverse=True)
            if ranked:
                chosen = ranked[0][1]

        if chosen is not None:
            chosen.set_focus()
            time.sleep(0.2)
            chosen.set_edit_text(export_name)
            filled = True
            LOGGER.info("Set File name to %s", export_name)

        send_keys("%s")  # Save
        time.sleep(0.6)
        confirm_replace_if_prompted(timeout=6.0)
        time.sleep(0.8)
    except Exception as exc:
        LOGGER.warning("Structured Save As failed (%s) — keyboard fallback", exc)
        filled = False

    if not filled:
        send_keys("%d")
        time.sleep(0.4)
        send_keys("^a")
        send_keys(str(export_dir), with_spaces=True)
        send_keys("{ENTER}")
        time.sleep(1.0)
        send_keys("%n")
        time.sleep(0.3)
        send_keys("^a")
        send_keys(export_name, with_spaces=True)
        send_keys("%s")
        time.sleep(0.6)
        confirm_replace_if_prompted(timeout=6.0)
        time.sleep(0.8)

    # Extra pass in case prompt appeared late
    confirm_replace_if_prompted(timeout=3.0)

    # If still no file, last-resort: put full path in filename box (works on this dialog)
    if not (target.exists() and target.stat().st_size > 0):
        still = None
        for w in Desktop(backend="win32").windows():
            if (w.window_text() or "").startswith("Save"):
                still = w
                break
        if still is not None:
            LOGGER.warning("Retry Save As using full path in File name as fallback")
            dlg = Application(backend="win32").connect(handle=still.handle).window(handle=still.handle)
            edits = list(dlg.descendants(class_name="Edit"))
            for edit in edits:
                text = edit.window_text() or ""
                if ".xls" in text.lower() or "report" in text.lower() or edit == edits[0]:
                    edit.set_edit_text(str(target))
                    break
            send_keys("%s")
            time.sleep(0.6)
            confirm_replace_if_prompted(timeout=6.0)
            time.sleep(0.8)

    deadline = time.time() + 60
    while time.time() < deadline:
        if target.exists() and target.stat().st_size > 0:
            LOGGER.info("Export saved: %s (%s bytes)", target, target.stat().st_size)
            return target
        matches = sorted(
            export_dir.glob(export_name.split(".")[0] + ".*"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if matches and matches[0].stat().st_mtime > time.time() - 120:
            LOGGER.info("Export saved (matched): %s", matches[0])
            return matches[0]
        time.sleep(1)

    raise TimeoutError(f"Export file not found at {target}")


def _csv_list(raw: str) -> list[str]:
    return [p.strip() for p in (raw or "").split(",") if p.strip()]


def configured_devices() -> list[str]:
    return _csv_list(os.getenv("ESSL_DEVICES", DEFAULT_DEVICES_CSV))


def configured_skip_devices() -> list[str]:
    return [d.upper() for d in _csv_list(os.getenv("ESSL_SKIP_DEVICES", DEFAULT_SKIP_CSV))]


def default_export_name(report: str | None = None) -> str:
    """Build Excel name like 'Aug Agra.xls' from LOCATION_CODE + current month."""
    code = (os.getenv("LOCATION_CODE") or "NOIDA").strip().upper()
    location = LOCATION_EXPORT_NAMES.get(code, code.title())
    month = datetime.now().strftime("%b")  # Jan, Feb, ... Aug
    return f"{month} {location}.xls"


def run(args: argparse.Namespace) -> int:
    load_dotenv()
    setup_logging()

    export_dir = Path(args.export_dir or os.getenv("ATTENDANCE_DIR", DEFAULT_EXPORT_DIR))
    user = args.user or os.getenv("ESSL_USER", DEFAULT_USER)
    password = args.password or os.getenv("ESSL_PASSWORD", DEFAULT_PASSWORD)
    devices = args.devices or configured_devices()
    skip_devices = configured_skip_devices()
    report = (args.report or DEFAULT_REPORT).lower()

    if args.dump_ui:
        dump_ui()
        return 0

    exe = find_exe()
    connect_app(exe)
    do_login(user, password)

    main = find_main_window(timeout=30)
    main.set_focus()
    time.sleep(0.5)

    if args.login_only:
        LOGGER.info("Login-only mode complete.")
        return 0

    if not args.skip_sync:
        open_device_management(main)
        select_devices(main, devices, skip_devices)
        if args.select_only:
            LOGGER.info("Select-only mode complete - inspect device checkboxes, then run without --select-only.")
            return 0
        # Always Start Download after selection, then wait until finished
        click_start_download(main)
        wait_download(main, args.download_timeout)
        if args.sync_only:
            LOGGER.info("Sync-only mode complete.")
            main = ensure_plain_ui(main)
            if not args.keep_open:
                do_logout(main)
                close_login_or_app()
            return 0
        main = find_main_window(timeout=10)

    # Always return to plain desktop before opening the attendance report menu
    main = ensure_plain_ui(main)
    open_report(main, report)
    generate_report(main, report=report)
    main = find_main_window(timeout=15)

    export_name = args.export_name or default_export_name(report)
    LOGGER.info("Export file name: %s (from LOCATION_CODE + month)", export_name)
    path = export_excel(main, export_dir, export_name)

    main = find_main_window(timeout=10)
    main = ensure_plain_ui(main)
    # Always log out after dialogs are closed (unless --keep-open)
    if args.keep_open:
        LOGGER.info("Skipping Log Off / Close because --keep-open was set")
    else:
        do_logout(main)
        close_login_or_app()

    LOGGER.info("DONE. File ready for importer: %s", path)
    LOGGER.info("Next: python scripts/importer_script.py")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="eSSL eTimeTrackLite pywinauto exporter")
    p.add_argument("--login-only", action="store_true", help="Login and stop")
    p.add_argument("--sync-only", action="store_true", help="Login + device download only")
    p.add_argument("--select-only", action="store_true", help="Select devices only (no Start Download)")
    p.add_argument("--skip-sync", action="store_true", help="Skip device download; export report only")
    p.add_argument("--dump-ui", action="store_true", help="Dump UIA tree and exit")
    p.add_argument(
        "--report",
        default=DEFAULT_REPORT,
        choices=["detailed", "basic", "monthly-basic"],
        help="Report type (default monthly-basic = Monthly Status / Basic Work Duration)",
    )
    p.add_argument("--keep-open", action="store_true", help="Do not Log Off after closing dialogs")
    p.add_argument("--logout", action="store_true", help="(deprecated) Log Off is default now")
    p.add_argument("--user", default=None, help="Login name")
    p.add_argument("--password", default=None, help="Password")
    p.add_argument("--export-dir", default=None, help="Excel output folder")
    p.add_argument("--export-name", default=None, help="Output file name only (not full path)")
    p.add_argument("--devices", nargs="+", default=None, help="Device names to keep selected")
    p.add_argument("--download-timeout", type=int, default=600, help="Max seconds to wait for device download")
    return p


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run(args)
    except Exception as exc:
        LOGGER.exception("essl_export failed: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
