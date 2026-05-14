# Code Structure + All Function/Method Definitions (Sequential)

## 1) Repository Structure
```text
CODE-/
├── gui/
│   ├── window1_entry.py
│   ├── window2_project_selection.py
│   ├── window3_scanner_selection.py
│   └── window4_main.py
├── logic/
│   ├── comparison_logic.py
│   ├── excel_writer
│   │   ├── __init__.py
│   │   ├── formatting.py
│   │   ├── reader.py
│   │   ├── sheets.py
│   │   ├── three_uk_qualys.py
│   │   ├── three_uk_qualys1.txt
│   │   ├── three_uk_qualys2.txt
│   │   ├── three_uk_qualys3.txt
│   │   └── workbook.py
│   ├── fast_vams_enrichment.py
│   ├── fast_vams_enrichment1.txt
│   ├── highlight_logic.py
│   ├── parser.py
│   ├── parser1.txt
│   ├── summary_generator.py
│   ├── vams_enrichment.py
│   └── vams_enrichment1.txt
├── tabs/
│   ├── add_vams_data.py
│   ├── add_vams_data1.txt
│   ├── add_vams_data_run_function.txt
│   ├── generate_tracking.py
│   └── make_new_report.py
├── utils/
│   ├── file_handler.py
│   ├── logger.py
│   └── memory.py
├── main.py
└── Requirements.txt
```

## 2) All `.py` definitions in sequence (file order)
### `gui/window1_entry.py`
- L7: `class Window1Entry(ttk.Frame):`
- L9: `def __init__(self, master, state, on_next):`
- L16: `def _build(self):`
- L22: `def _next(self):`

### `gui/window2_project_selection.py`
- L6: `class Window2ProjectSelection(ttk.Frame):`
- L8: `def __init__(self, master, state, on_prev, on_next):`
- L16: `def _build(self):`
- L26: `def _next(self):`

### `gui/window3_scanner_selection.py`
- L6: `class Window3ScannerSelection(ttk.Frame):`
- L8: `def __init__(self, master, state, on_prev, on_next):`
- L16: `def _build(self):`
- L26: `def _next(self):`

### `gui/window4_main.py`
- L11: `class Window4Main(ttk.Frame):`
- L13: `def __init__(self, master, state, logger):`
- L19: `def _build(self):`
- L30: `def attach_log_handler(self):`
- L37: `def append_log(self, msg: str):`

### `logic/comparison_logic.py`
- L13: `def classify_new_old(raw_df: pd.DataFrame, master_df: Optional[pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame]:`
- L24: `def aggregate_unique(df: pd.DataFrame) -> pd.DataFrame:`

### `logic/excel_writer/__init__.py`
- _(No class/def definitions)_

### `logic/excel_writer/formatting.py`
- L93: `def style_meta_row(`
- L115: `def write_headers(`
- L160: `def write_summary_headers(`
- L180: `def apply_table_formatting(ws):`
- L211: `def auto_width(ws):`

### `logic/excel_writer/reader.py`
- L22: `def _normalized_header(value: object) -> str:`
- L26: `def _fill_from_display_alias(df: pd.DataFrame, target_col: str):`
- L42: `def read_sheet_as_df(path, sheet_name):`

### `logic/excel_writer/sheets.py`
- L54: `def _write_data_rows(ws, df: pd.DataFrame):`
- L63: `def write_main_sheet(`
- L88: `def _clear_dashboard(ws):`
- L106: `def _hide_chart_source_columns(ws, start_col: int, width: int):`
- L117: `def _write_chart_source(ws, df: pd.DataFrame, start_row: int, start_col: int) -> int:`
- L127: `def _color_series_points(series, colors):`
- L135: `def _add_pie(ws, title: str, source_row: int, source_col: int, size: int, anchor: str):`
- L154: `def _add_bar(ws, title: str, source_row: int, source_col: int, rows: int, cols: int, anchor: str):`
- L175: `def _append_pie_charts(ws, total_df: pd.DataFrame, unique_df: pd.DataFrame, source_col: int = DASHBOARD_SOURCE_START_COL) -> int:`
- L189: `def _append_vams_bar_charts(ws, vams_df: pd.DataFrame, source_row: int, source_col: int = DASHBOARD_SOURCE_START_COL):`
- L218: `def write_summary_sheet(`
- L246: `def write_disposition_sheet(ws, project: str = "", scanner: str = "", *_, **__):`

### `logic/excel_writer/three_uk_qualys.py`
- L145: `def is_three_uk_qualys_project(`
- L166: `def detect_qualys_header_row(`
- L214: `def severity_to_criticality(`
- L256: `def criticality_series(`
- L281: `def map_risk(value):`
- L296: `def three_uk_qualys_total_view(`
- L355: `def build_3uk_qualys_total_sheet_df(`
- L382: `def build_3uk_vams_matching_df(`
- L417: `def build_3uk_qualys_unique_sheet_df(`
- L517: `def merge_semicolon_separated(`

### `logic/excel_writer/workbook.py`
- L43: `def _normalized_sheet_name(value: object) -> str:`
- L54: `def normalize_total_sheet_name(`
- L67: `def autofit_worksheet_columns(ws):`
- L102: `def write_output(`

### `logic/fast_vams_enrichment.py`
- L25: `def _norm(value: object) -> str:`
- L29: `def _split_multi_values(value):`
- L59: `def _extract_host_and_ports(value):`
- L118: `def _is_numeric_like(value):`
- L137: `def _is_empty(value: str) -> bool:`
- L147: `def build_fast_keys(`
- L370: `class FastVamsEnrichmentEngine:`
- L372: `def __init__(self):`
- L375: `def build_vams_lookup(`
- L410: `def enrich(`

### `logic/highlight_logic.py`
- L13: `def severity_fill(severity: str) -> PatternFill:`

### `logic/parser.py`
- L214: `class ParsedData:`
- L219: `def norm_text(value: object) -> str:`
- L229: `def split_values(value: object) -> List[str]:`
- L253: `def merge_semicolon(values: Iterable[object]) -> str:`
- L264: `def highest_risk(values: Iterable[object]) -> str:`
- L277: `def detect_header_row(path: str, sheet_name: str, scanner: str) -> int:`
- L320: `def _mapped_series(df: pd.DataFrame, target_col: str) -> pd.Series:`
- L340: `def normalize_columns(`
- L447: `def parse_scan_file(`
- L502: `def filter_severity(df: pd.DataFrame) -> pd.DataFrame:`
- L534: `def build_key(`

### `logic/summary_generator.py`
- L15: `def _normalized_series(df: pd.DataFrame, column: str) -> pd.Series:`
- L35: `def _severity_counts(df: pd.DataFrame) -> List[Dict[str, object]]:`
- L50: `def severity_summary(df: pd.DataFrame) -> pd.DataFrame:`
- L61: `def severity_chart_summary(df: pd.DataFrame, include_total: bool = True) -> pd.DataFrame:`
- L77: `def expert_severity_summary(df: pd.DataFrame) -> pd.DataFrame:`
- L93: `def _split_summary_values(value: object) -> List[str]:`
- L100: `def disposition_summary(df: pd.DataFrame, disposition_order: List[str]) -> pd.DataFrame:`

### `logic/vams_enrichment.py`
- L28: `def _norm(value: object) -> str:`
- L31: `def _split_multi_values(value):`
- L54: `def _extract_host_and_ports(value):`
- L97: `def _is_numeric_like(value):`
- L104: `def _norm_port(value: object) -> str:`
- L113: `def _fallback_cve_tokens(row: pd.Series) -> list[str]:`
- L148: `def _extract_port_tokens(value: object) -> list[str]:`
- L172: `def _extract_host_and_embedded_ports(`
- L211: `def _row_keys(`
- L264: `def _value_in_merged_cell(`
- L282: `def enrich_with_vams(`
- L426: `def update_vams_existing_workbook(`

### `main.py`
- L13: `class App(tk.Tk):`
- L15: `def __init__(self):`
- L30: `def _swap(self, frame):`
- L36: `def show_window1(self):`
- L39: `def show_window2(self):`
- L42: `def show_window3(self):`
- L45: `def show_window4(self):`

### `tabs/add_vams_data.py`
- L66: `class AddVamsDataTab(ttk.Frame):`
- L68: `def __init__(`
- L90: `def _build(self):`
- L205: `def _browse_output(self):`
- L214: `def _browse_raw(self):`
- L231: `def _validate_inputs(self):`
- L243: `def _read_generated_sheet(`
- L258: `def _write_vams_columns_only(`
- L425: `def _refresh_dashboard_charts(`
- L496: `def run(self):`

### `tabs/generate_tracking.py`
- L18: `class GenerateTrackingTab(ttk.Frame):`
- L20: `def __init__(self, master, app_state, logger):`
- L36: `def _build_ui(self):`
- L60: `def _file_row(self, row, label, var, browse_command, save=False):`
- L65: `def _load_sheets(self, path, combo, var):`
- L71: `def _browse_master(self):`
- L77: `def _browse_raw(self):`
- L83: `def _browse_vams(self):`
- L89: `def _browse_output(self):`
- L94: `def _validate_inputs(self):`
- L109: `def _parse_with_optional_severity_fallback(self, path: str, sheet: str, label: str):`
- L129: `def run(self):`

### `tabs/make_new_report.py`
- L22: `class MakeNewReportTab(ttk.Frame):`
- L24: `def __init__(self, master, app_state, logger):`
- L36: `def _build_ui(self):`
- L53: `def _browse_raw(self):`
- L63: `def _browse_output(self):`
- L68: `def _validate_inputs(self):`
- L75: `def run(self):`

### `utils/file_handler.py`
- L9: `def validate_file(path: str) -> None:`
- L17: `def list_excel_sheets(path: str) -> List[str]:`

### `utils/logger.py`
- L8: `class UILogHandler(logging.Handler):`
- L9: `def __init__(self, callback: Callable[[str], None]):`
- L13: `def emit(self, record: logging.LogRecord) -> None:`
- L18: `def get_logger(name: str = "vuln_automation", ui_callback: Optional[Callable[[str], None]] = None) -> logging.Logger:`

### `utils/memory.py`
- L11: `def memory_session(logger, label: str):`
- L30: `def release_large_objects(namespace: dict, names: list[str]):`
