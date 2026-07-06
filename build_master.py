"""
build_master.py — convert each EasyPower HTM report to an individual Excel file
and merge all PDFs into one.

Excel:
  Every SC*_DET.htm / SC*_SUM.htm report becomes its own .xlsx file
  (e.g. SC1_DET.htm -> SC1_DET.xlsx) in the output folder.

PDF:
  Every SC*.pdf is concatenated into Combined_Reports.pdf.

Run standalone (interactive — asks where the reports are):
    python build_master.py
or quick:
    python build_master.py <reports_folder>
It also runs automatically at the end of easypower_batch_reports.py.

Needs: pip install pandas openpyxl lxml pypdf
"""

import re
import sys
from pathlib import Path

import pandas as pd
from pypdf import PdfWriter


def _key(p):
    """Sort by scenario number then name: SC1_DET, SC1_SUM, SC2_DET, ... SC10..."""
    m = re.match(r"SC(\d+)", p.stem)
    return (int(m.group(1)) if m else 0, p.stem)


def _flatten_columns(df):
    """read_html turns multi-row table headers into MultiIndex columns, which
    pandas can't write to Excel with index=False. Join each column's levels into
    one name (dropping the empty/'Unnamed'/'nan' filler read_html inserts)."""
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(s for s in (str(x).strip() for x in col)
                     if s and s.lower() != "nan" and not s.startswith("Unnamed")).strip()
            for col in df.columns
        ]
    return df


def _explode_multiline_rows(df):
    """Expand rows where cells contain \\n (from <br> in the HTM) into separate
    rows.  EasyPower sometimes puts multiple branch entries under one bus into a
    single HTML table row separated by <br> tags.  read_html reads those as a
    single cell with newlines; we split them so each entry gets its own Excel row,
    leaving columns that didn't have breaks blank on subsequent rows."""
    br_cols = [col for col in df.columns
               if df[col].astype(str).str.contains("\n", na=False).any()]
    if not br_cols:
        return df

    new_rows = []
    for _, row in df.iterrows():
        parts = {}
        n = 1
        for col in br_cols:
            val = str(row[col]) if pd.notna(row[col]) else ""
            lines = val.split("\n")
            parts[col] = lines
            n = max(n, len(lines))
        for i in range(n):
            nr = {}
            for col in df.columns:
                if col in br_cols:
                    lines = parts[col]
                    nr[col] = lines[i].strip() if i < len(lines) else ""
                else:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    nr[col] = val.strip() if i == 0 else ""
            new_rows.append(nr)
    return pd.DataFrame(new_rows)


def htm_to_excels(folder, out_dir=None):
    """Convert each SC*_*.htm report into a separate .xlsx file.  Multiple
    <table> elements inside one HTM are stacked vertically on a single sheet."""
    if out_dir is None:
        out_dir = folder
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    htms = sorted(folder.glob("SC*_*.htm"), key=_key)
    if not htms:
        print("No SC*_*.htm reports found.")
        return

    converted = 0
    for f in htms:
        try:
            tables = pd.read_html(f)
        except Exception as e:
            print(f"skip {f.name}: {e}")
            continue

        out_path = out_dir / f"{f.stem}.xlsx"
        with pd.ExcelWriter(out_path, engine="openpyxl") as xl:
            sheet, row = f.stem[:31], 0
            for t in tables:
                t = _flatten_columns(t)
                t = _explode_multiline_rows(t)
                t.to_excel(xl, sheet_name=sheet, index=False, startrow=row)
                row += len(t) + 2
        converted += 1
        print(f"  {f.name} -> {out_path.name}")
    print(f"Converted {converted} report(s) to individual Excel files.")


def merge_pdfs(folder, out_pdf):
    """Every SC*.pdf -> one combined PDF, in scenario order."""
    pdfs = sorted(folder.glob("SC*.pdf"), key=_key)
    if not pdfs:
        print("No SC*.pdf files found.")
        return
    writer = PdfWriter()
    for f in pdfs:
        try:
            writer.append(str(f))
        except Exception as e:
            print(f"skip {f.name}: {e}")
    with open(out_pdf, "wb") as fh:
        writer.write(fh)
    print(f"Wrote {out_pdf} ({len(pdfs)} PDFs).")


def build(input_folder, output_folder=None):
    """Convert HTM reports to individual Excel files + merge PDFs."""
    in_dir = Path(input_folder)
    out_dir = Path(output_folder) if output_folder else in_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    htm_to_excels(in_dir, out_dir)
    merge_pdfs(in_dir, out_dir / "Combined_Reports.pdf")


def banner():
    """Styled title block (matches easypower_batch_reports.py). Duplicated here so
    this script doesn't have to import the pywinauto-heavy main module."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
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


if __name__ == "__main__":
    if len(sys.argv) > 1:                                # quick: folder on the command line
        build(sys.argv[1])
    else:
        banner()
        while True:
            src = input("Folder containing the report files: ").strip().strip('"')
            if Path(src).is_dir():
                break
            print("   That folder doesn't exist — try again.")
        dst = input(f"Output folder (for the .xlsx + .pdf files) [{src}]: ").strip().strip('"')
        build(src, dst or src)
