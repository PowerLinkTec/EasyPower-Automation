"""
build_master.py — combine per-scenario EasyPower report exports into ONE workbook.

Reads the intermediate files produced by easypower_batch_reports.py, named
    <scenario>__<reporttype>.xlsx   (or .csv)
and writes master.xlsx with:
  - one sheet per report type, every scenario stacked under a leading
    'Scenario' column (filter/pivot across all 36 in one place)
  - an 'Index' sheet showing which reports were found for each scenario,
    so any gaps from the batch run are obvious at a glance

Run after the UI batch run:
    python build_master.py <reports_folder> --out master.xlsx
"""

import re
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

NAME_RE = re.compile(r"^(?P<scenario>.+?)__(?P<report>.+)$")  # split on first '__'


def load_one(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


def collect(folder: Path):
    data, scenarios = {}, set()
    for f in sorted(folder.glob("*__*")):
        if f.suffix.lower() not in (".xlsx", ".csv"):
            continue
        m = NAME_RE.match(f.stem)
        if not m:
            continue
        scen, rpt = m["scenario"], m["report"]
        scenarios.add(scen)
        try:
            df = load_one(f)
        except Exception as e:
            print(f"skip {f.name}: {e}")
            continue
        df.insert(0, "Scenario", scen)
        data.setdefault(rpt, []).append((scen, df))
    return data, sorted(scenarios)


def _style(ws):
    fill = PatternFill("solid", start_color="1F4E78")
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial")
    for col in ws.columns:
        w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 2, 48)
    ws.freeze_panes = "B2"


def main():
    ap = argparse.ArgumentParser(description="Combine EasyPower report exports.")
    ap.add_argument("folder", type=Path, help="folder of <scenario>__<report> files")
    ap.add_argument("--out", type=Path, default=Path("master.xlsx"))
    args = ap.parse_args()

    data, scenarios = collect(args.folder)
    if not data:
        raise SystemExit(f"No '<scenario>__<report>.xlsx|csv' files in {args.folder}")

    report_types = sorted(data)
    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        idx = pd.DataFrame({"Scenario": scenarios})
        for rpt in report_types:
            present = {s for s, _ in data[rpt]}
            idx[rpt] = ["Y" if s in present else "" for s in scenarios]
        idx.to_excel(xl, sheet_name="Index", index=False)
        for rpt in report_types:
            combined = pd.concat([df for _, df in data[rpt]], ignore_index=True)
            combined.to_excel(xl, sheet_name=rpt[:31], index=False)

    wb = load_workbook(args.out)
    for ws in wb.worksheets:
        _style(ws)
    wb.save(args.out)
    print(f"Wrote {args.out}: Index + {len(report_types)} report sheet(s), "
          f"{len(scenarios)} scenario(s).")


if __name__ == "__main__":
    main()
