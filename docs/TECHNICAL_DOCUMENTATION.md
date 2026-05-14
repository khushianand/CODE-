# Vulnerability Management Automation - Full Technical Documentation

## 1. System Overview
This application is a desktop Tkinter workflow for vulnerability report processing:
- Parse scanner files and normalize columns.
- Compare raw data vs master baseline.
- Create new/old/unique datasets.
- Enrich using VAMS data.
- Write styled Excel workbooks with dashboards.

## 2. End-to-End Workflow
1. `main.py` launches `App` and stores state.
2. GUI wizard collects entry mode, project, and scanner.
3. Main window opens three tabs:
   - Tab1 Make New Report
   - Tab2 Update Tracking Sheet
   - Tab3 Add VAMS Data
4. Each tab validates files/sheets.
5. Parser converts source columns into canonical template.
6. Tab-specific computation executes.
7. Output workbook is written and formatted.
8. New memory helper logs and releases large objects after each run.

## 3. Core Logic and Computation
### 3.1 Parsing / Normalization (`logic/parser.py`)
- Header row detection scans preview rows and scores keyword-group overlap.
- Column mapping uses alias dictionary (`COLUMN_ALIASES`).
- Severity filters retain Critical/High/Medium/Low unless disabled.

Algorithm style:
- Heuristic scoring for header row detection.
- Dictionary-based alias lookup.

Complexity:
- Header detection: O(r * g) where r=preview rows, g=keyword groups.
- Column normalization: O(n * c) where n=rows, c=template columns.

### 3.2 New vs Old Classification (`logic/comparison_logic.py`)
- Builds deterministic comparison keys.
- Uses set membership to identify old rows efficiently.

DSA / algorithm:
- Hash set lookups.

Complexity:
- Key build: O(n)
- Membership checks: O(n) average

### 3.3 Unique Aggregation (`logic/comparison_logic.py`)
- Groups by (Name, CVE, Host / Image).
- Merges column values with deduped semicolon join.

DSA / algorithm:
- Hash grouping (`pandas groupby` semantics).
- Stable dedupe list for merged text values.

Complexity:
- Grouping and aggregation: near O(n) average (implementation-specific inside pandas).

### 3.4 VAMS Enrichment (`logic/vams_enrichment.py`, `logic/fast_vams_enrichment.py`)
- Builds multi-tier matching keys using host/name/cve/scanner/port.
- Applies non-destructive fill for VAMS columns.

DSA / algorithm:
- Tiered key matching with hash map lookup.
- Tokenization via split/regex for multi-value fields.

Complexity:
- Key generation: O(n * k) where k=number of generated key variants.
- Lookup/merge: O(n) average with hashing.

### 3.5 Excel Writer (`logic/excel_writer/*`)
- Writes workbook sheets, dashboard, and charts.
- Applies borders/alignment/row height/column width.

Algorithm style:
- Sequential worksheet traversal.

Complexity:
- Formatting cost proportional to number of written cells O(rows * cols).

## 4. Memory Behavior and Optimization
### 4.1 Why memory usage rises
- Large Excel files become in-memory pandas DataFrames.
- Some flows keep multiple DataFrames simultaneously (`raw_df`, `master_df`, `new_df`, `old_df`, `unique_df`, `vams_df`).
- Workbook objects can be large before close/save.

### 4.2 Implemented memory controls
Added `utils/memory.py`:
- `memory_session(...)` tracks current/peak memory for each tab execution.
- Force `gc.collect()` after task completion.
- `release_large_objects(...)` deletes known heavy variables and re-runs GC.

Applied in:
- Tab1 `tabs/make_new_report.py`
- Tab2 `tabs/generate_tracking.py`
- Tab3 `tabs/add_vams_data.py`

Behavior now:
- Memory is explicitly released when each tab run completes.
- Peak/current usage is logged so memory-heavy input patterns can be identified.

## 5. DSA and Algorithm Summary
- Hashing: key generation and fast membership lookup.
- Grouping/Aggregation: vulnerability dedupe and rollup.
- Heuristic scoring: robust header row detection.
- Tiered matching: progressive key fallback for enrichment accuracy.

## 6. File-by-File Responsibilities
- `main.py`: app bootstrap and window routing.
- `gui/window*.py`: wizard state capture and tab host.
- `tabs/make_new_report.py`: raw-only report generation.
- `tabs/generate_tracking.py`: raw/master comparison and optional VAMS enrichment.
- `tabs/add_vams_data.py`: enrich existing workbook with VAMS-only fields.
- `logic/parser.py`: canonical schema normalization.
- `logic/comparison_logic.py`: classify/aggregate logic.
- `logic/vams_enrichment.py`: merge rules and target VAMS columns.
- `logic/fast_vams_enrichment.py`: optimized tiered-key enrichment engine.
- `logic/excel_writer/*`: workbook creation, sheet writing, styling, dashboard charts.
- `utils/file_handler.py`: file and sheet helpers.
- `utils/logger.py`: app logging and UI log bridge.
- `utils/memory.py`: memory tracking and release utilities.

## 7. Documentation Format
The canonical documentation is maintained in this Markdown file for Git compatibility and easy review.
