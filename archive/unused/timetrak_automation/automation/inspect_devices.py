from pywinauto import Application, Desktop
import time

wins = [w for w in Desktop(backend="uia").windows() if "eTimeTrackLite" in (w.window_text() or "")]
print("windows:", [(w.window_text(), w.element_info.control_type) for w in wins])
if not wins:
    raise SystemExit("TimeTrak window not found")

app = Application(backend="uia").connect(handle=wins[0].handle)
win = app.window(handle=wins[0].handle)
win.set_focus()
time.sleep(0.5)

for c in win.descendants():
    try:
        name = c.window_text() or ""
        ctype = c.element_info.control_type
        if (
            ctype in ("DataGrid", "Table", "CheckBox", "List", "ListItem", "DataItem", "Custom")
            or "LGF" in name
            or "UGF" in name
            or "USB" in name
            or "Device" in name
        ):
            rect = c.rectangle()
            print(f"{ctype:10} | {name[:50]!r:52} | {rect}")
    except Exception:
        pass
