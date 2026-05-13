diff --git a/README.md b/README.md
index 6780f03e9b7a6b1e832f93623cbb68fe3797a8ff..2c06d14ed4c0f691edb296e0bb7a9b47b527035a 100644
--- a/README.md
+++ b/README.md
@@ -1 +1,193 @@
-# CODE-
\ No newline at end of file
+# Vulnerability Management Excel Automation
+
+This project is a desktop Tkinter application that automates vulnerability report generation and Excel workbook formatting for different scanner/project combinations.
+
+It provides a guided 4-step UI, then 3 processing tabs:
+- **Make New Report** (create a fresh workbook from a raw scan).
+- **Update Tracking Sheet** (compare raw vs master, classify new/old, optional VAMS enrichment).
+- **Add VAMS Data** (enrich an already generated workbook with VAMS fields).
+
+---
+
+## Requirements
+
+## Runtime
+- Python **3.10+** (recommended 3.11/3.12)
+- OS with GUI support (Tkinter available)
+
+## Python packages
+Install with:
+
+```bash
+pip install pandas openpyxl
+```
+
+Notes:
+- `tkinter` is used for the GUI and is usually bundled with Python distributions.
+- Excel I/O and formatting are handled via `openpyxl`.
+- Data transformation and aggregation are handled via `pandas`.
+
+---
+
+## How the application works (high-level)
+
+1. **App launch (`main.py`)**
+   - Creates root Tk window.
+   - Initializes shared state (`entry_mode`, `selected_project`, `selected_scanner`).
+   - Routes user through wizard windows (entry -> project -> scanner -> main tabs).
+
+2. **Wizard phase (`gui/`)**
+   - Window 1: choose mode (VNF/CNF).
+   - Window 2: choose project (PTA, 3UK, Airtel).
+   - Window 3: choose scanner (Nessus, Qualys, Anchor).
+   - Window 4: opens tabbed work area + live logs.
+
+3. **Execution phase (`tabs/`)**
+   - Each tab validates selected files/sheets.
+   - Raw/master/VAMS Excel files are parsed.
+   - Data is normalized and transformed.
+   - Outputs are written into standardized workbook sheets.
+   - Dashboard charts and formatting are applied.
+
+4. **Core business logic (`logic/`)**
+   - Parsing, severity filtering, normalization, dedupe/comparison, summary generation.
+   - Optional VAMS enrichment and fast key-based enrichment.
+   - Special handling for **3UK + Qualys** data layout.
+
+5. **Excel output engine (`logic/excel_writer/`)**
+   - Creates workbook and required sheets.
+   - Writes headers/data/dashboard/disposition.
+   - Applies table style, colors, borders, widths, charts.
+
+---
+
+## Workflow (end-to-end)
+
+1. Run app.
+2. Select entry mode -> project -> scanner.
+3. Open one of the 3 tabs:
+   - **Make New Report**: raw input + output path.
+   - **Update Tracking Sheet**: raw + optional master + optional VAMS + output path.
+   - **Add VAMS Data**: existing output workbook + raw VAMS workbook/sheet.
+4. Click **Run**.
+5. App validates input, parses sheets, applies comparison/enrichment logic, writes result workbook, formats workbook, and logs status in UI.
+
+---
+
+## How execution is triggered
+
+- Start command:
+
+```bash
+python main.py
+```
+
+- `if __name__ == "__main__": App().mainloop()` starts Tk event loop.
+- Button handlers in each tab call `run()` methods.
+- `run()` orchestrates parse -> transform -> write -> format -> notify.
+
+---
+
+## File-by-file explanation (`.py` files)
+
+## Root
+- **`main.py`**: Application entry point, root state holder, and window router for the 4-step GUI flow.
+
+## `gui/`
+- **`gui/window1_entry.py`**: Entry mode selection screen (VNF/CNF).
+- **`gui/window2_project_selection.py`**: Project selection screen.
+- **`gui/window3_scanner_selection.py`**: Scanner selection screen.
+- **`gui/window4_main.py`**: Main workspace with tabs + live log console; attaches UI log handler.
+
+## `tabs/`
+- **`tabs/make_new_report.py`**:
+  - Creates report from raw scan only.
+  - Builds total + unique vulnerability data.
+  - Writes output workbook (with dashboard).
+  - Applies final worksheet formatting.
+
+- **`tabs/generate_tracking.py`**:
+  - Tracking workflow: raw vs master comparison.
+  - Classifies new/old vulnerabilities.
+  - Aggregates unique findings.
+  - Optionally enriches with VAMS file.
+  - Writes formatted output workbook.
+
+- **`tabs/add_vams_data.py`**:
+  - Opens existing generated workbook.
+  - Reads target sheets (total/unique aliases supported).
+  - Parses VAMS raw source.
+  - Enriches workbook rows with VAMS columns using fast-key matching.
+  - Saves updated workbook and refreshes summary sheet/formatting.
+
+## `logic/`
+- **`logic/parser.py`**:
+  - Core parser/normalizer for scanner exports.
+  - Standardizes incoming columns to template columns.
+  - Applies severity filters and validation rules.
+
+- **`logic/comparison_logic.py`**:
+  - Compares raw vs master datasets.
+  - Produces `new_df` and `old_df`.
+  - Generates unique aggregation and VAMS join/enrichment helpers.
+
+- **`logic/vams_enrichment.py`**:
+  - Defines VAMS enrichment columns and merge logic.
+  - Adds VAMS attributes (disposition/comments/expert fields) to vulnerabilities.
+
+- **`logic/fast_vams_enrichment.py`**:
+  - Faster enrichment path using prebuilt keys/lookups for larger datasets.
+  - Used by Add VAMS tab for efficient row updates.
+
+- **`logic/summary_generator.py`**:
+  - Builds severity/disposition summary tables used by dashboard charts and counts.
+
+- **`logic/highlight_logic.py`**:
+  - Severity to Excel fill-color mapping for consistent risk highlighting.
+
+## `logic/excel_writer/`
+- **`logic/excel_writer/__init__.py`**:
+  - Public API exports for excel writer package.
+
+- **`logic/excel_writer/workbook.py`**:
+  - Workbook orchestration: create file, enforce sheet order, include dashboard/sheet options, save output.
+
+- **`logic/excel_writer/sheets.py`**:
+  - Writes main data sheets, summary/dashboard sheet, disposition sheet.
+  - Creates pie/bar charts for severity/disposition/expert metrics.
+
+- **`logic/excel_writer/reader.py`**:
+  - Reads generated workbook sheets into DataFrames.
+  - Uses display-column aliases to map headings back to template columns.
+
+- **`logic/excel_writer/formatting.py`**:
+  - Central style utilities: header styles, borders, alignments, row height, auto-width, table formatting.
+
+- **`logic/excel_writer/three_uk_qualys.py`**:
+  - Specialized transformation logic for the 3UK + Qualys export format.
+  - Detects header row and maps raw Qualys columns to required output schema.
+
+## `utils/`
+- **`utils/file_handler.py`**:
+  - Common file helpers (path validation, list Excel sheets, etc.).
+
+- **`utils/logger.py`**:
+  - Logger setup and UI logging bridge/handler used by the main window.
+
+---
+
+## Whole functioning summary
+
+The application is a **GUI-driven Excel automation pipeline**:
+- User chooses context (project + scanner).
+- App reads scan sheets into DataFrames.
+- Parser and business logic normalize and compare vulnerability records.
+- Optional VAMS data is joined for remediation/disposition enrichment.
+- Excel writer generates a standardized workbook with:
+  - dashboard,
+  - total/new/old/unique vulnerability sheets,
+  - optional disposition insights,
+  - formatting and charts.
+- Logs are streamed to the UI throughout execution.
+
+This keeps reporting consistent, repeatable, and much faster than manual spreadsheet operations.
