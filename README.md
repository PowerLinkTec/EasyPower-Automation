# EasyPower-Automation

Python scripts to run every scenario in an EasyPower study and export its Power Flow reports, automating a workflow EasyPower has no API for (so it drives the UI).

## Files

- **easypower_batch_reports.py** — the main script. Drives the EasyPower UI to loop over every scenario: open scenario → solve Power Flow → save the Detail and Summary reports (HTM) and a one-line PDF.
- **dez_inspect.py** — reads `.dez` files directly (they're ZIPs) to inventory each study's equipment/settings into a manifest, without opening EasyPower.
- **build_master.py** — combines the exported per-scenario report files into a single master Excel workbook.
- **ezp_inspect_ui.py** — utility that dumps EasyPower's live UI control tree, used to find the button/dialog identifiers the automation needs.
- **diag_menu.py** — one-off diagnostic for locating ribbon dropdown menu items (used to speed up the menu clicks).
