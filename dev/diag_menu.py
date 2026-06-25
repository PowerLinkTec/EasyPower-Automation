"""
diag_menu.py — locate an open ribbon dropdown's items so the batch script can
search them fast. Run with EasyPower open and a .dez loaded:

    python diag_menu.py "Scenario Mgr"

It opens that dropdown (in-process, so the menu stays open), then reports:
  - every top-level window in EasyPower's process (is the popup one of them?),
  - how long it takes to find MenuItems under the main window, and
  - the parent chain of each item (what container to scope the search to).
Paste the whole output back.
"""

import sys
import time
from pywinauto import Application, Desktop

BTN = sys.argv[1] if len(sys.argv) > 1 else "Scenario Mgr"
CLS = "EasyPowerClass"

app = Application(backend="uia").connect(class_name=CLS, timeout=10)
win = app.window(class_name=CLS)
pid = win.process_id()

# Open the dropdown via the same ribbon-scoped click the batch script uses.
rib = win.child_window(control_type="ToolBar", title="EasyPower", depth=1).wrapper_object()
for el in rib.descendants(control_type="SplitButton"):
    try:
        if " ".join((el.window_text() or "").split()) == BTN:
            el.click_input()
            break
    except Exception:
        pass
time.sleep(1.0)

print(f"\n=== top-level windows in EasyPower's process (pid {pid}) ===")
for w in Desktop(backend="uia").windows():
    try:
        if w.process_id() != pid:
            continue
        n = len(w.descendants(control_type="MenuItem")) if w.class_name() != CLS else "(skipped main)"
        print(f"class={w.class_name()!r}  ctype={w.element_info.control_type!r}  "
              f"title={w.window_text()!r}  menuitems={n}")
    except Exception as e:
        print(f"  (err: {e})")

print("\n=== MenuItems under the MAIN window ===")
t0 = time.time()
try:
    items = win.descendants(control_type="MenuItem")
except Exception as e:
    items = []
    print(f"  search failed: {e}")
print(f"win.descendants(MenuItem) -> {len(items)} items in {time.time() - t0:.1f}s")
for mi in items[:8]:
    chain, e = [], mi
    for _ in range(6):
        try:
            chain.append(f"{e.element_info.control_type}/{e.class_name()!r}")
            e = e.parent()
            if e is None:
                break
        except Exception:
            break
    print(f"  {mi.window_text()!r}  <-  " + " < ".join(chain))
