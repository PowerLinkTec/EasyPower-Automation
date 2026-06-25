"""
ezp_inspect_ui.py — dump EasyPower's live UI control names for wiring Stage 2.

Run with EasyPower OPEN (a study loaded, sitting where you'd normally start an
analysis). Writes ui_dump.txt: every window / button / menu / dialog field
pywinauto can see, with title, control_type and auto_id — the strings you
paste into the `# >>> ADAPT` spots of easypower_batch_reports.py.

    pip install pywinauto
    python ezp_inspect_ui.py --list              # show every open window's title
    python ezp_inspect_ui.py                      # dump the EasyPower window
    python ezp_inspect_ui.py --title "Base Case"  # match any part of the title
    python ezp_inspect_ui.py --backend win32      # if uia controls are sparse
    python ezp_inspect_ui.py --title "Save"       # dump an open dialog instead
    python ezp_inspect_ui.py --depth 6            # if the full dump is huge/slow

Matching is a plain case-insensitive substring of the title-bar text, using
the SAME enumeration that --list prints — so if a window shows in --list,
--title with part of its text will find it. (The old .window(title_re=...)
lookup matched the UIA Name property instead, which differs from the title bar
on MDI apps like EasyPower, so it never matched.)
ponytail: thin dumper, no logic to test — find a window and print it.
"""

import sys
import argparse
from pywinauto import Desktop

# Windows that merely MENTION a study (browser tabs, Explorer folders, the
# terminal running this script) aren't the app — skip them when listing/matching.
# ponytail: substring blocklist; add a term here if a new noise window sneaks in.
# Note: no bare "opera"/"edge" — they'd false-match "operating mode"/"knowledge".
EXCLUDE = (
    "file explorer", "chrome", "firefox", "microsoft edge",
    "google drive", "onedrive", "command prompt", "powershell",
    "windows terminal", "microsoft teams", "outlook", "program manager", "taskbar",
)


def is_noise(title):
    t = title.lower()
    return any(x in t for x in EXCLUDE)


def open_windows(backend):
    """Yield (title, wrapper) for every titled top-level window."""
    for w in Desktop(backend=backend).windows():
        try:
            t = w.window_text()
        except Exception:
            continue
        if t and t.strip():
            yield t, w


def list_windows():
    """Print every top-level window title under both backends, for discovery."""
    for backend in ("uia", "win32"):
        print(f"\n=== open windows ({backend}) ===")
        try:
            for t, _ in open_windows(backend):
                if not is_noise(t):
                    print(f"  {t!r}")
        except Exception as e:
            print(f"  (couldn't enumerate {backend}: {e})")


def dump_tree(e, fh, depth, max_depth, indent=0):
    """Recursively write one line per control: text, type, auto_id, class.
    Uses .children() — works on wrapper objects across pywinauto versions,
    unlike WindowSpecification.print_control_identifiers. Returns count."""
    info = getattr(e, "element_info", None)
    try:
        text = e.window_text()
    except Exception:
        text = ""
    ctype = getattr(info, "control_type", "") or ""
    autoid = getattr(info, "automation_id", "") or ""
    cls = getattr(info, "class_name", "") or ""
    fh.write(f"{'  ' * indent}text={text!r}  type={ctype!r}  "
             f"auto_id={autoid!r}  class={cls!r}\n")
    count = 1
    if max_depth is not None and depth >= max_depth:
        return count
    try:
        kids = e.children()
    except Exception:
        kids = []
    for k in kids:
        count += dump_tree(k, fh, depth + 1, max_depth, indent + 1)
    return count


def main():
    ap = argparse.ArgumentParser(description="Dump live UI control identifiers.")
    ap.add_argument("--title", default="EasyPower",
                    help="case-insensitive substring of the window title (default: EasyPower)")
    ap.add_argument("--backend", default="uia", choices=("uia", "win32"),
                    help="UI Automation backend (try win32 if uia controls are sparse)")
    ap.add_argument("--depth", type=int, default=None,
                    help="limit tree depth (try 6 if the dump is huge)")
    ap.add_argument("--out", default="ui_dump.txt")
    ap.add_argument("--list", action="store_true",
                    help="just list all open window titles and exit")
    args = ap.parse_args()

    if args.list:
        list_windows()
        return

    needle = args.title.lower()
    match = next(((t, w) for t, w in open_windows(args.backend)
                  if needle in t.lower() and not is_noise(t)), None)
    if match is None:
        print(f"No open window's title contains '{args.title}' (backend={args.backend}).\n"
              f"Here's what IS open — re-run --title with part of the right one "
              f"(or try --backend win32):")
        list_windows()
        sys.exit(1)

    title, win = match
    print(f"Dumping: {title!r}")
    try:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(f"# control tree for {title!r} (backend={args.backend})\n")
            n = dump_tree(win, fh, 0, args.depth)
    except Exception as e:
        print(f"Dump failed/partial ({e}). Try a smaller tree: --depth 6.")
        sys.exit(1)

    print(f"Wrote {args.out} ({n} controls). Open it, find the buttons/fields you "
          f"click by hand in EasyPower, copy their text / type / auto_id into the "
          f"`# >>> ADAPT` lines of easypower_batch_reports.py. Or send me the file.")


if __name__ == "__main__":
    main()
