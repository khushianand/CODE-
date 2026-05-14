# Full Project Documentation (Architecture, Logic, Workflow, Computation, DSA, Optimization)

## 1) Project Purpose
This project is a Tkinter desktop application for vulnerability Excel automation. It parses scan files, compares datasets, enriches with VAMS metadata, and generates formatted output workbooks with dashboard charts.

---

## 2) Whole-System Architecture
```mermaid
flowchart TB
    A[main.py App Router] --> B[GUI Wizard Windows]
    B --> C[window4_main.py Tabs Container]
    C --> D[tabs/make_new_report.py]
    C --> E[tabs/generate_tracking.py]
    C --> F[tabs/add_vams_data.py]

    D --> G[logic/parser.py]
    D --> H[logic/comparison_logic.py]
    D --> I[logic/excel_writer/*]

    E --> G
    E --> H
    E --> J[logic/vams_enrichment.py]
    E --> I

    F --> G
    F --> K[logic/fast_vams_enrichment.py]
    F --> I

    I --> L[logic/summary_generator.py]
    I --> M[logic/highlight_logic.py]

    D --> N[utils/file_handler.py]
    E --> N
    F --> N

    C --> O[utils/logger.py]
    D --> P[utils/memory.py]
    E --> P
    F --> P
```

---

## 3) Runtime Workflow (User + Compute)
```mermaid
flowchart LR
    U[User launches app] --> W1[Entry Mode]
    W1 --> W2[Project]
    W2 --> W3[Scanner]
    W3 --> W4[Tabs + Logs]
    W4 --> T1[Tab1 Make New Report]
    W4 --> T2[Tab2 Update Tracking]
    W4 --> T3[Tab3 Add VAMS]

    T1 --> X[Validate -> Parse -> Aggregate -> Write -> Format -> Cleanup]
    T2 --> Y[Validate -> Parse Raw/Master -> Classify -> Optional VAMS -> Write -> Format -> Cleanup]
    T3 --> Z[Validate -> Read Generated -> Parse VAMS -> Fast Enrich -> Writeback -> Refresh Dashboard -> Cleanup]
```

---

## 4) Imports and Dependency Map
- Standard library: `tkinter`, `logging`, `pathlib`, `re`, `gc`, `tracemalloc`, `dataclasses`, `typing`.
- Third-party: `pandas`, `openpyxl`.
- Internal packages:
  - `gui/*`
  - `tabs/*`
  - `logic/*`
  - `logic/excel_writer/*`
  - `utils/*`

---

## 5) Computation, DSA, Algorithms, Optimizations

### 5.1 Parsing & Normalization (`logic/parser.py`)
- **Algorithm**: header-row heuristic scoring + alias-based canonical mapping.
- **DSA**: dictionaries for alias lookup and normalization mapping.
- **Complexity**:
  - header detection: `O(r*g)` (preview rows × keyword groups)
  - normalization: approx `O(n*c)` (rows × target columns)
- **Optimization**: normalization aliases reduce scanner format coupling.

### 5.2 New/Old Classification (`logic/comparison_logic.py`)
- **Algorithm**: build deterministic keys and set-membership compare.
- **DSA**: hash-set membership.
- **Complexity**: near `O(n)` average membership checks.
- **Optimization**: avoids nested loops.

### 5.3 Unique Aggregation (`logic/comparison_logic.py`)
- **Algorithm**: group-by `(Name, CVE, Host / Image)` and de-dup merge.
- **DSA**: hash-grouping (`pandas groupby`) + stable merge lists.
- **Complexity**: near linear average in practice (pandas internals apply).

### 5.4 VAMS Enrichment (`logic/vams_enrichment.py`, `logic/fast_vams_enrichment.py`)
- **Algorithm**: tiered key matching with fallback combinations (host/name/cve/scanner/port).
- **DSA**: hash-map lookup + regex/token parsers.
- **Complexity**: `O(n*k)` key generation where `k` is tiered-key expansion.
- **Optimization**: fast-lookup precomputation avoids brute-force row-by-row nested matching.

### 5.5 Excel Generation (`logic/excel_writer/*`)
- **Algorithm**: sequential sheet writing + style passes + chart embedding.
- **DSA**: list/loop traversal.
- **Complexity**: formatting/writes roughly `O(rows*cols)`.
- **Optimization**: shared formatting helpers and reusable chart-source builders.

---

## 6) Memory Usage and Lifecycle

### Where memory is consumed most
1. `pandas.read_excel()` for raw/master/vams sheets.
2. Holding multiple DataFrames simultaneously (Tab2 highest).
3. `openpyxl` workbook object during write/save/format operations.

### What was added to control memory
- `utils/memory.py`
  - `memory_session(logger, label)` tracks `tracemalloc` current/peak.
  - `release_large_objects(...)` deletes heavy references and runs `gc.collect()`.
- Applied in all three tabs to force release after task completion.

### Approximate footprint behavior (depends on input size)
- Small file (<10k rows): tens of MB.
- Medium file (10k–100k rows): hundreds of MB possible.
- Large file (>100k rows with many columns): can exceed 1 GB when multiple DataFrames + workbook coexist.

---

## 7) Functional `def()` Flow and Flowcharts for All `.py` Files

## 7.1 `main.py`
- Class: `App`
- Flow: init state -> create container -> show wizard windows -> open tabbed workspace.
```mermaid
flowchart TD
    A[App.__init__] --> B[show_window1]
    B --> C[show_window2]
    C --> D[show_window3]
    D --> E[show_window4]
    E --> F[attach_log_handler]
```

## 7.2 `gui/window1_entry.py`
- Class: `Window1Entry`
- Methods: `_build`, `_next`
```mermaid
flowchart TD
    A[_build] --> B[User selects VNF/CNF]
    B --> C[_next updates state]
    C --> D[Go to window2]
```

## 7.3 `gui/window2_project_selection.py`
- Class: `Window2ProjectSelection`
- Methods: `_build`, `_next`
```mermaid
flowchart TD
    A[_build] --> B[User selects project]
    B --> C[_next updates selected_project]
    C --> D[Go to window3]
```

## 7.4 `gui/window3_scanner_selection.py`
- Class: `Window3ScannerSelection`
- Methods: `_build`, `_next`
```mermaid
flowchart TD
    A[_build] --> B[User selects scanner]
    B --> C[_next updates selected_scanner]
    C --> D[Go to window4]
```

## 7.5 `gui/window4_main.py`
- Class: `Window4Main`
- Methods: `_build`, `attach_log_handler`, `append_log`
```mermaid
flowchart TD
    A[_build notebook tabs] --> B[MakeNewReportTab]
    A --> C[GenerateTrackingTab]
    A --> D[AddVamsDataTab]
    E[attach_log_handler] --> F[UILogHandler]
    F --> G[append_log]
```

## 7.6 `tabs/make_new_report.py`
- Class: `MakeNewReportTab`
- Core method: `run()`
```mermaid
flowchart TD
    A[run] --> B[validate inputs]
    B --> C{3UK + Qualys?}
    C -->|Yes| D[build_3uk total + unique]
    C -->|No| E[parse_scan_file + aggregate_unique]
    D --> F[write_output]
    E --> F
    F --> G[load_workbook + apply_table_formatting]
    G --> H[save + success]
    H --> I[memory cleanup]
```

## 7.7 `tabs/generate_tracking.py`
- Class: `GenerateTrackingTab`
- Core method: `run()`
```mermaid
flowchart TD
    A[run] --> B[validate]
    B --> C[parse raw]
    C --> D[optional parse master]
    D --> E[classify_new_old]
    E --> F[aggregate unique]
    F --> G{vams file?}
    G -->|Yes| H[parse vams + enrich_with_vams]
    G -->|No| I[skip]
    H --> J[write_output]
    I --> J
    J --> K[format workbook]
    K --> L[memory cleanup]
```

## 7.8 `tabs/add_vams_data.py`
- Class: `AddVamsDataTab`
- Key methods: `_read_generated_sheet`, `_write_vams_columns_only`, `_refresh_dashboard_charts`, `run`
```mermaid
flowchart TD
    A[run] --> B[validate + load generated sheets]
    B --> C[parse incoming VAMS]
    C --> D[build FastVamsEnrichmentEngine lookup]
    D --> E[enrich unique]
    E --> F{3UK + Qualys?}
    F -->|No| G[enrich+write total and unique]
    F -->|Yes| H[write unique only]
    G --> I[refresh dashboard]
    H --> I
    I --> J[memory cleanup]
```

## 7.9 `utils/file_handler.py`
- Functions:
  - `validate_file(path)`
  - `list_excel_sheets(path)`
```mermaid
flowchart TD
    A[validate_file] --> B[path exists/readable checks]
    C[list_excel_sheets] --> D[open workbook metadata + sheet names]
```

## 7.10 `utils/logger.py`
- Class: `UILogHandler`
- Function: `get_logger(...)`
```mermaid
flowchart TD
    A[get_logger] --> B[configure logger and formatter]
    B --> C{UI callback given?}
    C -->|Yes| D[attach UILogHandler]
    C -->|No| E[console logger only]
```

## 7.11 `utils/memory.py`
- Functions:
  - `memory_session(logger, label)`
  - `release_large_objects(namespace, names)`
```mermaid
flowchart TD
    A[memory_session enter] --> B[start tracemalloc if needed]
    B --> C[run task]
    C --> D[gc.collect + log end/peak]
    E[release_large_objects] --> F[del references]
    F --> G[gc.collect]
```

## 7.12 `logic/parser.py`
- Main functions/classes: `ParsedData`, `norm_text`, `split_values`, `merge_semicolon`, `highest_risk`, `detect_header_row`, `_mapped_series`, `normalize_columns`, `parse_scan_file`, `filter_severity`, `build_key`.
```mermaid
flowchart TD
    A[parse_scan_file] --> B[detect_header_row]
    B --> C[read excel with header]
    C --> D[normalize_columns]
    D --> E[filter_severity optional]
    E --> F[return ParsedData]
```

## 7.13 `logic/comparison_logic.py`
- Functions: `classify_new_old`, `aggregate_unique`
```mermaid
flowchart TD
    A[classify_new_old] --> B[build key sets]
    B --> C[split into new_df / old_df]
    D[aggregate_unique] --> E[groupby Name,CVE,Host]
    E --> F[merge_semicolon + highest_risk]
```

## 7.14 `logic/vams_enrichment.py`
- Functions: normalization helpers, token extraction, key construction, `enrich_with_vams`, `update_vams_existing_workbook`.
```mermaid
flowchart TD
    A[enrich_with_vams] --> B[build VAMS lookup keys]
    B --> C[build target row keys]
    C --> D[tiered match]
    D --> E[fill VAMS columns non-destructively]
```

## 7.15 `logic/fast_vams_enrichment.py`
- Functions: key-building helpers + `build_fast_keys`.
- Class: `FastVamsEnrichmentEngine` (`build_vams_lookup`, `enrich`, internal match).
```mermaid
flowchart TD
    A["build_vams_lookup"]
    B["precompute hash maps by tiers"]
    C["enrich(df)"]
    D["build_fast_keys per row"]
    E["lookup tiered keys"]
    F["merge VAMS columns"]

    A --> B
    C --> D
    D --> E
    E --> F
```

## 7.16 `logic/summary_generator.py`
- Functions: severity and disposition summaries.
```mermaid
flowchart TD
    A[severity_summary/chart_summary] --> B[count normalized severities]
    C[expert_severity_summary] --> D[compare reported vs expert]
    E[disposition_summary] --> F[tokenize and count disposition labels]
```

## 7.17 `logic/highlight_logic.py`
- Function: `severity_fill`.
```mermaid
flowchart TD
    A[severity_fill] --> B[map severity to color]
    B --> C[return PatternFill]
```

## 7.18 `logic/excel_writer/formatting.py`
- Functions: `style_meta_row`, `write_headers`, `write_summary_headers`, `apply_table_formatting`, `auto_width`.
```mermaid
flowchart TD
    A[write_headers] --> B[meta row styles]
    B --> C[column header styles]
    D[apply_table_formatting] --> E[borders + alignment + row heights]
    F[auto_width] --> G[column sizing pass]
```

## 7.19 `logic/excel_writer/reader.py`
- Functions: `_normalized_header`, `_fill_from_display_alias`, `read_sheet_as_df`.
```mermaid
flowchart TD
    A[read_sheet_as_df] --> B[read header row 2]
    B --> C[fill target fields using aliases]
    C --> D[return TEMPLATE_COLUMNS dataframe]
```

## 7.20 `logic/excel_writer/workbook.py`
- Functions: `_normalized_sheet_name`, `normalize_total_sheet_name`, `autofit_worksheet_columns`, `write_output`.
```mermaid
flowchart TD
    A[write_output] --> B[create workbook]
    B --> C{include dashboard?}
    C -->|Yes| D[write_summary_sheet]
    C -->|No| E[first data sheet]
    D --> F[write new/old/unique/disposition]
    E --> F
    F --> G[save workbook]
```

## 7.21 `logic/excel_writer/sheets.py`
- Functions for data rows, dashboard chart data, pie/bar chart creation, sheet writers.
```mermaid
flowchart TD
    A[write_main_sheet] --> B[headers + rows + freeze panes + width]
    C[write_summary_sheet] --> D[summary tables]
    D --> E[pie/bar charts]
    F[write_disposition_sheet] --> G[disposition view]
```

## 7.22 `logic/excel_writer/three_uk_qualys.py`
- Functions for 3UK+Qualys detection, header detection, mapping, total/unique builders.
```mermaid
flowchart TD
    A[build_3uk_qualys_total_sheet_df] --> B[detect_qualys_header_row]
    B --> C[map raw qualys fields]
    C --> D[criticality/risk transformations]
    D --> E[return standardized total DF]
    F[build_3uk_qualys_unique_sheet_df] --> G[group/map into unique template]
```

## 7.23 `logic/excel_writer/__init__.py`
- Re-exports public writer APIs for simple imports.

---

## 8) Practical Optimization Recommendations (Next Steps)
1. Use `usecols=` in `pd.read_excel` where feasible.
2. Convert high-cardinality string columns to categoricals for large workflows.
3. Process large files in stages/chunks when possible.
4. Skip global formatting passes on sheets that are unchanged.
5. Add benchmark scripts for row-count vs memory/latency tracking.
6. Add test suite for parser alias correctness and enrichment key quality.

---

## 9) Summary
This project uses a layered architecture (GUI -> Tabs -> Logic -> Excel Writer -> Utilities), hash-based comparison/enrichment logic, and explicit memory lifecycle cleanup. It is optimized for operational report automation with configurable scanner/project-specific normalization and reusable formatting/chart generation.
