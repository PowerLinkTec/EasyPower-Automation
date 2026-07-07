# EasyPower-Automation

Python scripts to run every scenario in an EasyPower study and export its Power Flow reports, automating a workflow EasyPower has no API for (so it drives the UI).

**New here?** See **[SETUP.md](SETUP.md)** for step-by-step setup from scratch (installing Python, etc.).

## Files

- **easypower_batch_reports.py** — the main script. Drives the EasyPower UI to loop over every scenario: open scenario → solve Power Flow → save the Detail and Summary reports (HTM) and a one-line PDF.
- **build_master.py** — post-processes a finished run: converts the HTM reports into one Excel workbook (a sheet per report) and merges the one-line PDFs into one combined PDF. Runs automatically when the batch finishes, or standalone.
- **powerflow-report.py** — takes the per-scenario detail/summary Excel files produced by `build_master.py` (`sc_XX_det.xlsx`, `sc_XX_sum.xlsx`) and fills their power-flow results (voltages, currents, power, power factor) into the **PF_Master.xlsx** template for IFC submission. Outputs `PF_Master_IFC.xlsx`. Run standalone after `build_master.py`.
  - **PF_Master.xlsx** — a hand-authored template (4 sheets: scenario list, power, voltages, branch loading) that `powerflow-report.py` populates with simulation results from each scenario. Tracked in git; edit the template directly when columns or layout change.

`dev/` — helper/inspection tools used while building this; not needed for normal use:
- **dev/dez_inspect.py** — reads `.dez` files directly (they're ZIPs) to inventory a study's equipment/settings, without opening EasyPower.
- **dev/ezp_inspect_ui.py** — dumps EasyPower's live UI control tree, used to find the button/dialog identifiers the automation needs.
- **dev/diag_menu.py** — diagnostic for locating ribbon dropdown menu items (used to speed up the menu clicks).
