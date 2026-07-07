"""
EasyPower Power Flow batch report extractor
===========================================
EasyPower has no scripting API, so this drives the UI via pywinauto. For every
scenario inside the .dez that is ALREADY OPEN in EasyPower, it repeats your
manual workflow:

    1. Scenario Mgr -> Open Scenario -> (store? Yes) -> select the scenario
    2. Power Flow -> Solve
    3. (make the Detail/Summary text-report windows exist)
     4. Window -> Power Flow Detail Report  -> File>Save As -> sc_<n>_det.htm
     5. close that report (Ctrl+F4)
     6. Window -> Power Flow Summary Report -> File>Save As -> sc_<n>_sum.htm  (+close)
     7. Print -> OK -> Save Print Output As -> sc_<n>.pdf  (Microsoft Print to PDF)
    8. Database Edit  (back to edit mode, ready for the next scenario)

TWO BACKENDS: EasyPower's ribbon/menus/backstage are modern (uia), but its
classic dialogs (Open Scenario, Print One-line, the Yes/No message boxes) are
Win32/MFC and INVISIBLE to uia. So we attach both: `win` (uia) for the ribbon,
`app32` (win32) for those dialogs. Modern file pickers stay on uia.

Spots that still need a live check are marked `# >>> VERIFY`.

Requirements: Windows; EasyPower running with the target .dez open; a held
Bentley license; `pip install pywinauto`. Run a COPY of the .dez if you don't
want stored-scenario results written back to the original.
"""

import re
import sys
import time
import logging
from pathlib import Path

from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.timings import Timings

Timings.window_find_timeout = 20     # give the (slow) whole-window fallback room

# --------------------------------------------------------------------------- CONFIG
OUTPUT_DIR = Path(r"C:\studies\reports")     # default; the startup prompt overrides it
MAIN_WINDOW_CLASS = "EasyPowerClass"         # match on class, not title (the title
                                             # collides with Explorer/Chrome windows)
SETTLE = 0.3          # pause after each UI action settles (tune down if stable)
SOLVE_WAIT = 3.0      # seconds to let a Solve finish (no API event exists)

log = logging.getLogger("ezp")


# --------------------------------------------------------------------------- helpers
def attach():
    """Attach BOTH backends to the running EasyPower (with the .dez open)."""
    try:
        app = Application(backend="uia").connect(class_name=MAIN_WINDOW_CLASS, timeout=10)
        app32 = Application(backend="win32").connect(class_name=MAIN_WINDOW_CLASS, timeout=10)
    except Exception:
        sys.exit("EasyPower not found. Start it and open the .dez first.")
    win = app.window(class_name=MAIN_WINDOW_CLASS)
    win.wait("visible", timeout=60)
    return app, win, app32


_RIBBON = None       # cached ribbon-container wrapper (stable; its children are
                     # queried live each call, so no per-button staleness)


def _ribbon(win):
    """Cached ribbon container. Searching only its subtree avoids walking the
    huge one-line canvas, which is what makes whole-window uia searches ~10s.
    The ribbon is a direct child of the main window (depth=1 = fast to locate)."""
    global _RIBBON
    if _RIBBON is None:
        _RIBBON = win.child_window(control_type="ToolBar", title="EasyPower",
                                   depth=1).wrapper_object()
    return _RIBBON


def _norm(s):
    return " ".join(s.split())   # collapse newlines/runs of spaces in button text


def _keys_escape(s):
    """Escape characters special to send_keys so a literal path/text is typed."""
    s = s.replace("{", "{{}").replace("}", "{}}")
    for ch in "+^%~()[]":
        s = s.replace(ch, "{" + ch + "}")
    return s


def click(win, title, ctype=None):
    """uia click by visible text. Fast path: scan only the ribbon subtree (skips
    the slow one-line canvas). Fallback: whole-window search for controls outside
    the ribbon (dropdown menu items, backstage). Resolved FRESH every call so the
    element and its coordinates are always current -- caching button wrappers made
    later scenarios click stale ribbon buttons (Power Flow / Solve were missed)."""
    want, el = _norm(title), None
    try:
        for cand in _ribbon(win).descendants(control_type=ctype):
            try:
                if _norm(cand.window_text()) == want:
                    el = cand
                    break
            except Exception:
                continue
    except Exception:
        globals()["_RIBBON"] = None          # ribbon wrapper went stale; re-resolve
    if el is None:                            # not in ribbon: dropdown / backstage
        kw = {"title": title}
        if ctype:
            kw["control_type"] = ctype
        el = win.child_window(**kw).wrapper_object()
    el.click_input()
    time.sleep(SETTLE)


def w32_button(dlg, text):
    """win32: click a dialog button by text (BM_CLICK via .click() — coordinate-
    free, so DPI scaling can't make it miss). Tolerates an & accelerator."""
    dlg.child_window(title_re=f"^&?{re.escape(text)}$", class_name="Button").click()
    time.sleep(SETTLE)


def click_menu_item(win, pattern):
    """Click an item in the open ribbon dropdown. Confirmed via diag: the dropdown
    is a 'Menu' element that is a DIRECT child of the main window. So find it with
    win.children() (direct children only -- instant) and search just that tiny
    subtree. This avoids win.descendants(MenuItem), which crawls the one-line
    canvas and measured ~28s. Falls back to that slow search only if not found."""
    rx = re.compile(pattern)
    for _ in range(10):                                  # menu may take a beat to render
        try:
            menus = win.children(control_type="Menu")
        except Exception:
            menus = []
        for menu in menus:
            try:
                for mi in menu.descendants(control_type="MenuItem"):
                    if rx.search(mi.window_text() or ""):
                        mi.click_input()
                        time.sleep(SETTLE)
                        return
            except Exception:
                continue
        time.sleep(0.2)
    win.child_window(title_re=f".*{pattern}.*", control_type="MenuItem").click_input()
    time.sleep(SETTLE)


def dismiss_store_prompt(app32, store=True):
    """Win32 message box: 'The scenario has changed. Do you want to store it?'
    Yes stores it (your choice). Harmless if it never appears."""
    try:
        dlg = app32.window(title="EasyPower", class_name="#32770")
        dlg.wait("visible", timeout=2)
        w32_button(dlg, "Yes" if store else "No")
    except Exception:
        pass


def dismiss_pf_options(app32, wait=0):
    """OK the 'Power Flow Options' dialog if it's open. It can pop up at more than
    one point in the Power Flow -> report path (after Solve, or when a report is
    generated), so this is called defensively around those steps. `wait` = seconds
    to allow it to appear; 0 = instant check (cheap). Harmless if absent."""
    try:
        dlg = app32.window(title="Power Flow Options")
        if dlg.exists(timeout=wait):
            w32_button(dlg, "OK")
            return True
    except Exception:
        pass
    return False


def _confirm_overwrite(app32):
    """Win32 'file exists, replace?' prompt -> Yes. No-op if absent."""
    try:
        dlg = app32.window(class_name="#32770")
        dlg.wait("visible", timeout=2)
        w32_button(dlg, "Yes")
    except Exception:
        pass


def scenario_num(name):
    """'Scenario-10' -> '10'. Falls back to a filename-safe form of the name."""
    m = re.search(r"(\d+)\s*$", name)
    return m.group(1).zfill(2) if m else re.sub(r"[^A-Za-z0-9]+", "_", name)


# --------------------------------------------------------------------------- steps
def open_scenario_dialog(win, app32):
    """uia: Scenario Mgr > Open Scenario ; returns the win32 Open Scenario dialog."""
    click(win, "Scenario Mgr", "SplitButton")
    click_menu_item(win, "Open Scenario")           # >>> VERIFY exact text
    dlg = app32.window(title="Open Scenario")             # classic dialog -> win32
    dlg.wait("visible", timeout=15)
    return dlg


def list_scenarios(win, app32):
    """Open the dialog, read the scenario list (win32 ListBox), Cancel out."""
    click(win, "Scenario Mgr", "SplitButton")
    click_menu_item(win, "Open Scenario")
    dismiss_store_prompt(app32, store=False)             # don't store just to peek
    dlg = app32.window(title="Open Scenario")
    dlg.wait("visible", timeout=15)
    lb = dlg.child_window(class_name="ListBox")
    names = lb.item_texts() if hasattr(lb, "item_texts") else lb.texts()[1:]
    w32_button(dlg, "Cancel")
    return [n for n in names if n]


def open_scenario(win, app32, name):
    """Step 1."""
    click(win, "Scenario Mgr", "SplitButton")
    click_menu_item(win, "Open Scenario")
    dismiss_store_prompt(app32, store=True)              # store current, per your flow
    dlg = app32.window(title="Open Scenario")
    dlg.wait("visible", timeout=15)
    dlg.child_window(class_name="ListBox").select(name)  # win32 select-by-text
    w32_button(dlg, "Open")
    try:                                                 # confirm it actually opened
        dlg.wait_not("visible", timeout=6)
    except Exception:
        log.info("Open didn't take; retrying %s", name)
        dlg.child_window(class_name="ListBox").select(name)
        w32_button(dlg, "Open")
    log.info("Opened %s", name)


def solve_power_flow(win, app32):
    """Step 2: Power Flow (switch focus) -> Solve. The 'Power Flow Options' dialog
    pops up AFTER Solve, so wait for it and OK it before moving on."""
    click(win, "Power Flow", "SplitButton")              # shows the POWER FLOW tab
    dismiss_pf_options(app32)                            # instant check
    click(win, "Solve", "Button")                        # Action group; not Solve Motor
    dismiss_pf_options(app32, wait=6)                    # usually appears here -> OK it
    time.sleep(SOLVE_WAIT)                               # let the solve finish


def save_report(win, app, app32, report_name, out_path):
    """Steps 4/6: Window > <report> -> click the report's Save button -> type
    the full output path into the Save Report dialog -> Save -> close the report.
    The full path in 'File name' makes Windows save to OUTPUT_DIR regardless of the
    dialog's current folder. 'Save Report' is a standard file dialog (uia)."""
    click(win, "Window", "SplitButton")
    dismiss_pf_options(app32)                            # instant: in case it blocks
    click_menu_item(win, re.escape(report_name))         # fast popup-scoped search
    dismiss_pf_options(app32, wait=4)                    # generating the report may pop it
    time.sleep(SETTLE)
    try:                                                 # report toolbar Save button
        click(win, "Save", "Button")                     # -> opens the Save Report dialog
    except Exception:
        win.type_keys("^s")                              # fallback: Save shortcut
    dlg = app32.window(title="Save Report")              # this dialog is win32, not uia
    dlg.wait("visible", timeout=15)
    dlg.set_focus()
    time.sleep(SETTLE)
    # Standard Windows save dialog: the File name field has focus on open. Select
    # all, type the full path, Enter = Save (the dump shows the default button is
    # '&Save'). Avoids the unlabeled Edit-inside-ComboBox that selectors can't hit.
    send_keys("^a" + _keys_escape(str(out_path)) + "{ENTER}", with_spaces=True, pause=0.02)
    _confirm_overwrite(app32)                            # 'replace existing?' -> Yes
    time.sleep(SETTLE)
    win.type_keys("^{F4}")                                # close the report window
    time.sleep(SETTLE)
    log.info("Saved %s", out_path.name)


def print_to_pdf(win, app, app32, out_pdf):
    """Step 7: Print -> OK (Microsoft Print to PDF) -> Save Print Output As.
    'Print One-line' is a classic dialog (win32); the PDF picker is modern (uia)."""
    click(win, "Print", "SplitButton")                   # QAT print, top-left
    dlg = app32.window(title="Print One-line")           # classic dialog -> win32
    dlg.wait("visible", timeout=15)
    w32_button(dlg, "OK")
    save = app32.window(title="Save Print Output As")    # standard Windows dialog (win32)
    save.wait("visible", timeout=30)
    save.set_focus()
    time.sleep(SETTLE)
    send_keys("^a" + _keys_escape(str(out_pdf)) + "{ENTER}", with_spaces=True, pause=0.02)
    _confirm_overwrite(app32)
    time.sleep(SETTLE)
    log.info("Printed %s", out_pdf.name)


def database_edit(win):
    """Step 8: back to edit mode, ready for the next scenario."""
    click(win, "Database Edit", "Button")


# --------------------------------------------------------------------------- run
def process_scenario(win, app, app32, name):
    n = scenario_num(name)
    prefix = f"sc_{n}"
    outs = [OUTPUT_DIR / f"{prefix}_det.htm",
            OUTPUT_DIR / f"{prefix}_sum.htm",
            OUTPUT_DIR / f"{prefix}.pdf"]
    if all(o.exists() for o in outs):                    # resume support
        log.info("Skip %s (all outputs exist)", name)
        return
    open_scenario(win, app32, name)
    solve_power_flow(win, app32)
    save_report(win, app, app32, "Power Flow Detail Report",  outs[0])
    save_report(win, app, app32, "Power Flow Summary Report", outs[1])
    print_to_pdf(win, app, app32, outs[2])
    database_edit(win)


def banner():
    """Print the styled title block."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")     # so the box + © render on Windows
    except Exception:
        pass
    lines = [
        "EasyPower Reporting Automation",
        "Copyright © Power-Link Technologies 2026",
        "",
        "Reach out to cooper@powerlinktec.com",
        "for troubleshooting and for reporting bugs",
    ]
    w = max(len(s) for s in lines) + 6
    print("\n╔" + "═" * w + "╗")
    for s in lines:
        print("║" + s.center(w) + "║")
    print("╚" + "═" * w + "╝\n")


def _confirm(prompt):
    """Block until the user types yes (anything else re-asks)."""
    while input(prompt).strip().lower() not in ("y", "yes"):
        print("   ...type 'yes' when ready.")


def prompt_setup():
    """Interactive startup: collect the output dir and scenario list, then gate the
    run on the user confirming scenario naming, EasyPower is open, and settings
    are configured.  Returns (output_dir, only, excludes) where `only` is a list
    of scenario names or None for all, and `excludes` are names to filter out."""
    banner()
    raw = input(f"Output directory [{OUTPUT_DIR}]: ").strip().strip('"')
    output_dir = Path(raw) if raw else OUTPUT_DIR

    raw = input("Scenarios to run (comma-separated names, or blank for ALL; '-' prefix to exclude): ").strip()
    parts = [s.strip() for s in raw.split(",") if s.strip()]
    includes = [p for p in parts if not p.startswith("-")]
    excludes = [p[1:] for p in parts if p.startswith("-")]
    only = includes or None

    print()
    _confirm("Ensure each scenario name ends with a number (e.g. Scenario-1), then type yes: ")
    _confirm("Open EasyPower and load the Base Case .dez, then type yes: ")
    _confirm("Set your printing settings, then type yes: ")
    _confirm("Set your report settings, then type yes: ")
    print("\n" + "!" * 64)
    print("  STARTING RUN — do NOT touch the mouse or keyboard until it")
    print("  finishes. The script is controlling the screen for you.")
    print("!" * 64 + "\n")
    return output_dir, only, excludes


def main():
    global OUTPUT_DIR
    OUTPUT_DIR, only, excludes = prompt_setup()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        handlers=[logging.StreamHandler(),
                  logging.FileHandler(OUTPUT_DIR / "batch_run.log", encoding="utf-8")],
    )
    app, win, app32 = attach()
    scenarios = only or list_scenarios(win, app32)
    if excludes:
        scenarios = [s for s in scenarios if s not in excludes]
    log.info("Found %d scenarios.", len(scenarios))
    failures = []
    for i, name in enumerate(scenarios, 1):
        log.info("[%d/%d] %s", i, len(scenarios), name)
        try:
            process_scenario(win, app, app32, name)
        except Exception as e:
            log.exception("FAILED on %s: %s", name, e)
            failures.append(name)
            database_edit(win)                           # try to get back to a clean state
    log.info("Done. %d ok, %d failed.", len(scenarios) - len(failures), len(failures))
    if failures:
        log.warning("Failed: %s", ", ".join(failures))

    log.info("Converting reports to individual Excel files + one PDF...")
    try:
        from build_master import build
        build(OUTPUT_DIR)
    except Exception as e:
        log.warning("Combine step failed (%s). Raw files are still in %s.", e, OUTPUT_DIR)

    log.info("Populating PF Master.xlsx with power-flow results...")
    try:
        import importlib.util
        _script_dir = Path(__file__).parent
        spec = importlib.util.spec_from_file_location(
            "powerflow_report", _script_dir / "powerflow-report.py"
        )
        pfr = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(pfr)
        det_dir = OUTPUT_DIR / "ez_automation_output"
        pfr.build(det_dir, det_dir)
    except Exception as e:
        log.warning("PowerFlow report step failed (%s).", e)


if __name__ == "__main__":
    main()
