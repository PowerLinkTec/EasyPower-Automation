"""
build_master.py — convert each EasyPower HTM report to an individual Excel file,
merge one-line diagram PDFs, and produce a combined PDF printout of all tables.

HTM reports:
  Every SC*_DET.htm / SC*_SUM.htm and LF_*_DET.htm / LF_*_SUM.htm report becomes
  its own .xlsx file (e.g. SC1_DET.htm -> SC1_DET.xlsx) in the output folder.

PDF:
  One-line diagrams (SLDs): Every SC*.pdf / LF_*.pdf is concatenated into
  combined_sld.pdf (scenario order, LF after SC).
  Data tables: Every SC*_*.xlsx / LF_*_*.xlsx is rendered into combined_report.pdf
  via reportlab (SC first, then LF by percentage).

Run standalone (interactive — asks where the reports are):
    python build_master.py
or quick:
    python build_master.py <reports_folder>
It also runs automatically at the end of easypower_batch_reports.py.

Needs: pip install pandas openpyxl lxml pypdf reportlab
"""

import re
import sys
from pathlib import Path

import numpy as np

# NumPy 2.x removed deprecated aliases like np.float, np.int, np.bool.
# Some libraries (e.g. openpyxl with older pandas) still reference them.
for _name, _val in (("float", float), ("int", int), ("bool", bool),
                     ("complex", complex)):
    if not hasattr(np, _name):
        setattr(np, _name, _val)

import pandas as pd

from pypdf import PdfWriter

# Patch pypdf's md5 call – some OpenSSL builds reject 'usedforsecurity'.
import pypdf._writer, hashlib
_orig = pypdf._writer._rolling_checksum
def _safe_rolling_checksum(stream, blocksize=65536):
    h = hashlib.md5()
    for block in iter(lambda: stream.read(blocksize), b""):
        h.update(block)
    return h.hexdigest()
pypdf._writer._rolling_checksum = _safe_rolling_checksum


def _key(p):
    """Sort: SC files first by number, then LF files by percentage.  Files that
    match neither pattern sort at the end."""
    m_sc = re.match(r"SC(\d+)", p.stem)
    m_lf = re.match(r"LF_(\d+)", p.stem)
    if m_sc:
        return (0, int(m_sc.group(1)), p.stem)
    if m_lf:
        return (1, int(m_lf.group(1)), p.stem)
    return (2, 0, p.stem)


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
    """Convert each SC*_*.htm / LF_*_*.htm report into a separate .xlsx file.
    Multiple <table> elements inside one HTM are stacked vertically on a single
    sheet."""
    if out_dir is None:
        out_dir = folder
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    htms = sorted(list(folder.glob("SC*_*.htm")) + list(folder.glob("LF_*_*.htm")), key=_key)
    if not htms:
        print("No HTM reports found.")
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
    """Merge one-line diagrams (SLDs): every SC*.pdf / LF_*.pdf -> one combined PDF."""
    pdfs = sorted(list(folder.glob("SC*.pdf")) + list(folder.glob("LF_*.pdf")), key=_key)
    if not pdfs:
        print("No PDF files found for merging.")
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


def xlsx_to_combined_pdf(folder, out_pdf):
    """Convert every SC*_*.xlsx / LF_*_*.xlsx into one combined PDF via reportlab.

    Each sheet from each workbook becomes a section in the PDF with a bold
    heading and a table rendered in 7pt Helvetica with grid lines, light-grey
    header background, and automatic page splitting.  SC files appear first,
    then LF files in percentage order."""
    xlsx_files = sorted(list(folder.glob("SC*_*.xlsx")) + list(folder.glob("LF_*_*.xlsx")), key=_key)
    if not xlsx_files:
        print("No Excel files found for PDF printout.")
        return

    from openpyxl import load_workbook
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph,
    )
    from reportlab.lib.styles import getSampleStyleSheet

    doc = SimpleDocTemplate(
        str(out_pdf), pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=10*mm, bottomMargin=10*mm,
    )

    styles = getSampleStyleSheet()
    heading_style = styles["Heading2"]
    elements = []

    for f in xlsx_files:
        wb = load_workbook(f, read_only=True, data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            elements.append(
                Paragraph(f"<b>{f.stem}</b> &mdash; {sheet_name}", heading_style)
            )

            data = []
            for row in ws.iter_rows(values_only=True):
                data.append([str(v) if v is not None else "" for v in row])

            if data:
                t = Table(data, repeatRows=1, hAlign="LEFT")
                t.setStyle(
                    TableStyle([
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                        ("FONTSIZE", (0, 0), (-1, -1), 7),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.Color(0.6, 0.6, 0.6)),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("TOPPADDING", (0, 0), (-1, -1), 1),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                        ("LEFTPADDING", (0, 0), (-1, -1), 2),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ])
                )
                elements.append(t)
            elements.append(PageBreak())
        wb.close()

    # Drop trailing blank page
    if elements and isinstance(elements[-1], PageBreak):
        elements.pop()

    if elements:
        doc.build(elements)
        print(f"Wrote {out_pdf} ({len(xlsx_files)} Excel files).")
    else:
        print("No data to write to PDF.")


def build(input_folder, output_folder=None):
    """Convert HTM reports to individual Excel files + merge PDFs + combined report."""
    in_dir = Path(input_folder)
    out_dir = Path(output_folder) if output_folder else in_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    htm_to_excels(in_dir, out_dir)
    merge_pdfs(in_dir, out_dir / "combined_sld.pdf")
    xlsx_to_combined_pdf(out_dir, out_dir / "combined_report.pdf")


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
