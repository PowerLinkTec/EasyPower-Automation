"""
build_master.py — combine the EasyPower batch outputs in a folder into two
deliverables:
  - Combined_Reports.xlsx : every SC*_DET.htm / SC*_SUM.htm report converted to a
    sheet (one report per sheet, in scenario order);
  - Combined_Reports.pdf  : every SC*.pdf one-line print merged into one PDF.

Run standalone (interactive — asks where the reports are and where the combined
files should go):
    python build_master.py
or quick, writing the combined files into the same folder:
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


def htm_to_workbook(folder, out_xlsx):
    """Each SC*_*.htm report -> one sheet. All HTML tables in a report are stacked
    so nothing is dropped (relies on pandas' startrow-append, fine on modern pandas)."""
    htms = sorted(folder.glob("SC*_*.htm"), key=_key)
    if not htms:
        print("No SC*_*.htm reports found.")
        return
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xl:
        for f in htms:
            try:
                tables = pd.read_html(f)
            except Exception as e:                       # no tables / no parser
                print(f"skip {f.name}: {e}")
                continue
            sheet, row = f.stem[:31], 0
            for t in tables:
                _flatten_columns(t).to_excel(xl, sheet_name=sheet, index=False, startrow=row)
                row += len(t) + 2
    print(f"Wrote {out_xlsx} ({len(htms)} reports).")


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
    """Convert + merge the reports in input_folder; write the two combined files
    to output_folder (defaults to input_folder)."""
    in_dir = Path(input_folder)
    out_dir = Path(output_folder) if output_folder else in_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    htm_to_workbook(in_dir, out_dir / "Combined_Reports.xlsx")
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
        dst = input(f"Folder for the combined files [{src}]: ").strip().strip('"')
        build(src, dst or src)
