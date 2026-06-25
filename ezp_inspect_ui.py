"""
ezp_inspect_ui.py — dump EasyPower's live UI control names for wiring Stage 2.

Run this with EasyPower already OPEN (a study loaded, sitting where you'd
normally start an analysis). It writes ui_dump.txt listing every window /
button / menu / dialog field pywinauto can see, with their title,
control_type and auto_id — the exact strings you paste into the `# >>> ADAPT`
spots of easypower_batch_reports.py.

    pip install pywinauto
    python ezp_inspect_ui.py --list           # show every open window's title
    python ezp_inspect_ui.py                   # dump the EasyPower window
    python ezp_inspect_ui.py --title "PV&BESS" # match by project name instead
    python ezp_inspect_ui.py --backend win32   # if uia can't see the window
    python ezp_inspect_ui.py --title "Save"    # dump an open dialog
    python ezp_inspect_ui.py --depth 6         # if the full dump is huge/slow

If --title finds nothing, it prints every open window so you can see the real
title to match. ponytail: thin dumper, no logic to test — connect and print.
"""

import sys
import argparse
from pywinauto import Desktop


def list_windows():
    """Print every top-level window title under both backends, for discovery."""
    for backend in ("uia", "win32"):
        print(f"\n=== open windows ({backend}) ===")
        try:
            for w in Desktop(backend=backend).windows():
                try:
                    t = w.window_text()
                except Exception:
                    t = ""
                if t and t.strip():
                    print(f"  {t!r}")
        except Exception as e:
            print(f"  (couldn't enumerate {backend}: {e})")


def main():
    ap = argparse.ArgumentParser(description="Dump live UI control identifiers.")
    ap.add_argument("--title", default="EasyPower",
                    help="substring of the window/dialog title (default: EasyPower)")
    ap.add_argument("--backend", default="uia", choices=("uia", "win32"),
                    help="UI Automation backend (try win32 if uia can't see it)")
    ap.add_argument("--depth", type=int, default=None,
                    help="limit tree depth (try 6 if the dump is huge)")
    ap.add_argument("--out", default="ui_dump.txt")
    ap.add_argument("--list", action="store_true",
                    help="just list all open window titles and exit")
    args = ap.parse_args()

    if args.list:
        list_windows()
        return

    try:
        win = Desktop(backend=args.backend).window(title_re=f".*{args.title}.*")
        win.wait("exists", timeout=10)
    except Exception:
        print(f"No window matching '{args.title}' (backend={args.backend}).\n"
              f"Here's what IS open — re-run with --title set to part of the right "
              f"one (add --backend win32 if EasyPower only appears in that list):")
        list_windows()
        sys.exit(1)

    with open(args.out, "w", encoding="utf-8") as fh:
        old, sys.stdout = sys.stdout, fh
        try:
            win.print_control_identifiers(depth=args.depth)
        finally:
            sys.stdout = old

    print(f"Wrote {args.out}. Open it, find the buttons/fields you click by hand "
          f"in EasyPower, and copy their title / control_type / auto_id into the "
          f"`# >>> ADAPT` lines of easypower_batch_reports.py.")


if __name__ == "__main__":
    main()
