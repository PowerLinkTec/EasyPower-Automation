"""
EasyPower batch report extractor
=================================
Opens each EasyPower file in a folder, runs a set of analyses, and exports
each report to disk. EasyPower has no scripting/COM API, so this drives the
Windows UI via pywinauto (Windows UI Automation backend).

Requirements
------------
- Windows, with EasyPower installed and a Bentley license seat available.
- An interactive, logged-in desktop session (this cannot run truly headless).
- pip install pywinauto

How to adapt (do this once)
---------------------------
The pieces that depend on your EasyPower version are marked `# >>> ADAPT`.
To discover the right menu/dialog identifiers, open EasyPower, then in a
Python shell:

    from pywinauto import Application
    app = Application(backend="uia").connect(title_re=".*EasyPower.*")
    app.window(title_re=".*EasyPower.*").print_control_identifiers()

or use Microsoft's "Accessibility Insights" / Inspect.exe to read control
names. Paste those names into the helpers below.
"""

import sys
import time
import logging
from pathlib import Path

from pywinauto import Application
from pywinauto.timings import wait_until, TimeoutError as PWTimeout

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
EASYPOWER_EXE = r"C:\Program Files\EasyPower\EasyPower.exe"   # >>> ADAPT path
INPUT_DIR     = Path(r"C:\studies\scenarios")                # the 36 files live here
OUTPUT_DIR    = Path(r"C:\studies\reports")
FILE_GLOB     = "*.dez"                                       # >>> ADAPT extension if different

MAIN_WINDOW_RE = ".*EasyPower.*"
CALC_TIMEOUT   = 600        # seconds to allow one analysis to finish
SETTLE         = 1.0        # small pause after UI actions

# Each report is a recipe the loop runs for every file. Add/remove freely.
# `run` and `export` reference helper funcs you fill in below.
REPORTS = [
    {"name": "arcflash",     "focus": "Arc Flash"},        # >>> ADAPT focus/menu labels
    {"name": "shortcircuit", "focus": "Short Circuit"},
    {"name": "coordination", "focus": "Coordination"},
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    handlers=[logging.StreamHandler(),
              logging.FileHandler(OUTPUT_DIR / "batch_run.log", encoding="utf-8")],
)
log = logging.getLogger("ezp")


# ---------------------------------------------------------------------------
# Low-level UI helpers  (the version-specific surface area)
# ---------------------------------------------------------------------------
def launch():
    """Start EasyPower (or attach if already running) and return the app/window."""
    try:
        app = Application(backend="uia").connect(title_re=MAIN_WINDOW_RE, timeout=5)
        log.info("Attached to running EasyPower instance.")
    except Exception:
        app = Application(backend="uia").start(EASYPOWER_EXE)
        log.info("Launched EasyPower.")
    win = app.window(title_re=MAIN_WINDOW_RE)
    win.wait("visible ready", timeout=120)
    return app, win


def open_file(app, win, path: Path):
    """
    Open a study file. Try the command-line/Recent route first; fall back to
    File > Open. Many Windows apps accept a path argument, which is far less
    brittle than clicking through the Open dialog.
    """
    # Preferred: relaunch-with-arg style is unreliable for an already-open app,
    # so use the Open dialog. >>> ADAPT the menu path + dialog field names.
    win.menu_select("File->Open")          # may be ribbon, not classic menu — adapt
    time.sleep(SETTLE)
    dlg = app.window(title_re="Open.*")
    dlg.child_window(auto_id="1148", control_type="Edit").set_text(str(path))  # filename box
    dlg.child_window(title="Open", control_type="Button").click()
    win.wait("visible ready", timeout=120)
    time.sleep(SETTLE)
    log.info("Opened %s", path.name)


def select_focus(win, focus_label: str):
    """Switch EasyPower to the analysis 'focus' for this report. >>> ADAPT."""
    win.menu_select(f"Home->{focus_label}")   # ribbon/menu label per report recipe
    time.sleep(SETTLE)


def run_analysis(win):
    """Trigger the calculation and block until it completes. >>> ADAPT trigger."""
    win.menu_select("Home->Run")              # or a toolbar button; adapt
    wait_for_calc(win)


def wait_for_calc(win):
    """
    No API => poll the UI until the calc is done. Adapt the predicate to
    something reliable in your version, e.g. a progress dialog disappearing
    or a results status control appearing.
    """
    def done():
        # >>> ADAPT: return True when results are ready.
        # Example: a progress window is gone, or a "Done"/status control exists.
        try:
            return not win.child_window(title_re="Calculating.*").exists()
        except Exception:
            return True
    try:
        wait_until(CALC_TIMEOUT, 1.0, done, value=True)
    except PWTimeout:
        raise RuntimeError("Calculation did not finish within timeout.")
    time.sleep(SETTLE)


def export_report(app, win, report_name: str, out_path: Path):
    """
    Open the report and export it. Arc Flash / Equipment Duty reports export to
    Excel or CSV from the report's spreadsheet view; text reports export/print
    to file. >>> ADAPT the open-report + export-button + save-dialog steps.
    """
    win.menu_select("Reports->Create Report")     # adapt to the report type
    time.sleep(SETTLE)
    rpt = app.window(title_re=".*Report.*")
    rpt.child_window(title="Excel", control_type="Button").click()   # or CSV/PDF
    time.sleep(SETTLE)
    save = app.window(title_re="Save As.*")
    save.child_window(auto_id="1001", control_type="Edit").set_text(str(out_path))
    save.child_window(title="Save", control_type="Button").click()
    # handle a possible "overwrite?" confirmation
    try:
        app.window(title_re="Confirm.*").child_window(title="Yes").click()
    except Exception:
        pass
    time.sleep(SETTLE)
    log.info("Exported %s", out_path.name)


def close_file(win):
    """Close the current file without saving. >>> ADAPT."""
    try:
        win.menu_select("File->Close")
        time.sleep(SETTLE)
        # decline any save prompt
        win.type_keys("{ENTER}")   # adapt: explicitly click "No"/"Don't Save"
    except Exception as e:
        log.warning("Close issue: %s", e)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def out_name(study: Path, report_name: str) -> Path:
    return OUTPUT_DIR / f"{study.stem}__{report_name}.xlsx"


def process_file(app, win, study: Path):
    open_file(app, win, study)
    for rpt in REPORTS:
        target = out_name(study, rpt["name"])
        if target.exists():                      # resume support: skip done work
            log.info("Skip (exists): %s", target.name)
            continue
        select_focus(win, rpt["focus"])
        run_analysis(win)
        export_report(app, win, rpt["name"], target)
    close_file(win)


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(INPUT_DIR.glob(FILE_GLOB))
    if not files:
        log.error("No files matching %s in %s", FILE_GLOB, INPUT_DIR)
        sys.exit(1)
    log.info("Found %d files; %d reports each.", len(files), len(REPORTS))

    app, win = launch()
    failures = []
    for i, study in enumerate(files, 1):
        log.info("[%d/%d] %s", i, len(files), study.name)
        try:
            process_file(app, win, study)
        except Exception as e:
            log.exception("FAILED on %s: %s", study.name, e)
            failures.append(study.name)
            # try to get back to a clean state for the next file
            try:
                close_file(win)
            except Exception:
                pass

    log.info("Done. %d ok, %d failed.", len(files) - len(failures), len(failures))
    if failures:
        log.warning("Failed files: %s", ", ".join(failures))


if __name__ == "__main__":
    main()
