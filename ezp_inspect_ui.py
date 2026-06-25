"""
ezp_inspect_ui.py — dump EasyPower's live UI control names for wiring Stage 2.

Run this with EasyPower already OPEN (a study loaded, sitting where you'd
normally start an analysis). It writes ui_dump.txt listing every window /
button / menu / dialog field pywinauto can see, with their title,
control_type and auto_id — the exact strings you paste into the `# >>> ADAPT`
spots of easypower_batch_reports.py.

    pip install pywinauto
    python ezp_inspect_ui.py                  # dump the main EasyPower window
    python ezp_inspect_ui.py --title "Save"   # dump an open dialog instead
    python ezp_inspect_ui.py --depth 6        # if the full dump is huge/slow

To capture a transient dialog (File>Open, Save As, the report window): get it
on screen first and leave it open, then run with --title matching its title
bar. ponytail: thin dumper, no logic to test — it connects and prints.
"""

import sys
import argparse
from pywinauto import Desktop


def main():
    ap = argparse.ArgumentParser(description="Dump live UI control identifiers.")
    ap.add_argument("--title", default="EasyPower",
                    help="substring of the window/dialog title (default: EasyPower)")
    ap.add_argument("--depth", type=int, default=None,
                    help="limit tree depth (try 6 if the dump is huge)")
    ap.add_argument("--out", default="ui_dump.txt")
    args = ap.parse_args()

    try:
        win = Desktop(backend="uia").window(title_re=f".*{args.title}.*")
        win.wait("exists", timeout=10)
    except Exception as e:
        sys.exit(f"No window matching '{args.title}'. Is it open and visible? ({e})")

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
