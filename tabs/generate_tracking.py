"""Tab 2: Update tracking sheet using optional master + optional VAMS enrichment."""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from logic.comparison_logic import aggregate_unique, classify_new_old
from logic.excel_writer import write_output
from logic.excel_writer.formatting import (apply_table_formatting,)
from logic.parser import parse_scan_file
from logic.excel_writer.three_uk_qualys import (
    build_3uk_qualys_total_sheet_df,
    build_3uk_qualys_unique_sheet_df,
)
from utils.file_handler import list_excel_sheets, validate_file
from utils.memory import memory_session, release_large_objects
from utils.performance import optimize_dataframe_memory, should_skip_full_formatting


class GenerateTrackingTab(ttk.Frame):
    """Core workflow: parse -> compare -> aggregate -> optional VAMS -> write output."""
    def __init__(self, master, app_state, logger):
        """Explain workflow and purpose of `__init__` in this module."""
        super().__init__(master)
        self.state = app_state
        self.logger = logger

        self.master_file = tk.StringVar()
        self.master_sheet = tk.StringVar()
        self.raw_file = tk.StringVar()
        self.raw_sheet = tk.StringVar()
        self.output_file = tk.StringVar()


        self._build_ui()

    def _build_ui(self):
        """Explain workflow and purpose of `_build_ui` in this module."""
        self._file_row(0, "Master File (optional)", self.master_file, self._browse_master)
        ttk.Label(self, text="Master sheet").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.master_combo = ttk.Combobox(self, textvariable=self.master_sheet, state="readonly", width=35)
        self.master_combo.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        self._file_row(2, "Raw File", self.raw_file, self._browse_raw)
        ttk.Label(self, text="Raw sheet").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.raw_combo = ttk.Combobox(self, textvariable=self.raw_sheet, state="readonly", width=35)
        self.raw_combo.grid(row=3, column=1, sticky="w", padx=6, pady=6)

        self._file_row(4, "Output File", self.output_file, self._browse_output, save=True)


        self.run_btn = ttk.Button(self, text="Run", command=self.run)
        self.run_btn.grid(row=5, column=1, sticky="e", padx=6, pady=10)
        self.columnconfigure(1, weight=1)


    def _file_row(self, row, label, var, browse_command, save=False):
        """Explain workflow and purpose of `_file_row` in this module."""
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(self, textvariable=var, width=72).grid(row=row, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(self, text="Browse", command=browse_command).grid(row=row, column=2, padx=6, pady=6)

    def _load_sheets(self, path, combo, var):
        """Explain workflow and purpose of `_load_sheets` in this module."""
        sheets = list_excel_sheets(path)
        combo["values"] = sheets
        if sheets:
            var.set(sheets[0])

    def _browse_master(self):
        """Explain workflow and purpose of `_browse_master` in this module."""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if path:
            self.master_file.set(path)
            self._load_sheets(path, self.master_combo, self.master_sheet)

    def _browse_raw(self):
        """Explain workflow and purpose of `_browse_raw` in this module."""
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if path:
            self.raw_file.set(path)
            self._load_sheets(path, self.raw_combo, self.raw_sheet)

    def _browse_output(self):
        """Explain workflow and purpose of `_browse_output` in this module."""
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.output_file.set(path)

    def _validate_inputs(self):
        """Explain workflow and purpose of `_validate_inputs` in this module."""
        validate_file(self.raw_file.get())
        if not self.raw_sheet.get():
            raise ValueError("Please select raw sheet")
        if self.master_file.get():
            validate_file(self.master_file.get())
            if not self.master_sheet.get():
                raise ValueError("Please select master sheet")
        if not self.output_file.get():
            raise ValueError("Please select output file")

    def _parse_with_optional_severity_fallback(self, path: str, sheet: str, label: str):
        """Explain workflow and purpose of `_parse_with_optional_severity_fallback` in this module."""
        try:
            return parse_scan_file(path, sheet, self.state["selected_project"], self.state["selected_scanner"]).df
        except ValueError as exc:
            if "No rows left after severity filtering" not in str(exc):
                raise
            self.logger.warning(
                "%s has no rows after severity filtering; retrying without severity filter",
                label,
            )
            return parse_scan_file(
                path,
                sheet,
                self.state["selected_project"],
                self.state["selected_scanner"],
                require_rows_after_filter=False,
                apply_severity_filter=False,
            ).df


    def run(self):

        """Explain workflow and purpose of `run` in this module."""
        try:
        
            self.run_btn.configure(
                state="disabled"
            )
            mem_ctx = memory_session(self.logger, "TAB2 Update Tracking Sheet")
            mem_ctx.__enter__()
    
            self._validate_inputs()
    
            # -------------------------------------------------
            # PARSE RAW FILE
            # -------------------------------------------------
    
            if (
                self.state["selected_project"].strip().casefold() == "3uk"
                and self.state["selected_scanner"].strip().casefold() == "qualys"
            ):

                raw_df = optimize_dataframe_memory(build_3uk_qualys_total_sheet_df(
                    self.raw_file.get(),
                    self.raw_sheet.get(),
                ))

            else:
            
                raw_df = optimize_dataframe_memory(self._parse_with_optional_severity_fallback(
                    self.raw_file.get(),
                    self.raw_sheet.get(),
                    "Raw file",
                ))
    
            # -------------------------------------------------
            # PARSE MASTER FILE
            # -------------------------------------------------
    
            master_df = None

            if self.master_file.get():
            
                if (
                    self.state["selected_project"].strip().casefold() == "3uk"
                    and self.state["selected_scanner"].strip().casefold() == "qualys"
                ):

                    master_df = optimize_dataframe_memory(build_3uk_qualys_total_sheet_df(
                        self.master_file.get(),
                        self.master_sheet.get(),
                    ))

                else:
                
                    master_df = optimize_dataframe_memory(self._parse_with_optional_severity_fallback(
                        self.master_file.get(),
                        self.master_sheet.get(),
                        "Master file",
                    ))
            # -------------------------------------------------
            # CLASSIFICATION
            # -------------------------------------------------
    
            new_df, old_df = classify_new_old(
                raw_df,
                master_df,
            )
    
            if (
                self.state["selected_project"].strip().casefold() == "3uk"
                and self.state["selected_scanner"].strip().casefold() == "qualys"
            ):
            
                unique_df = optimize_dataframe_memory(build_3uk_qualys_unique_sheet_df(
                    self.raw_file.get(),
                    self.raw_sheet.get(),
                ))
            
            else:
            
                unique_df = optimize_dataframe_memory(aggregate_unique(raw_df))
                    
    
            # -------------------------------------------------
            # WRITE OUTPUT
            # -------------------------------------------------
    
            output = write_output(
                self.output_file.get(),
                new_df,
                old_df,
                unique_df,
                self.state["selected_project"],

                self.state["selected_scanner"],
            )
    
            # -------------------------------------------------
            # APPLY PROFESSIONAL FORMATTING
            # -------------------------------------------------
    
            from openpyxl import load_workbook
    
            from logic.excel_writer.formatting import (
                apply_table_formatting,
            )
    
            if should_skip_full_formatting(raw_df, unique_df):
                self.logger.warning("Large dataset detected; skipping full worksheet formatting for performance.")
            else:
                wb = load_workbook(output)
    
                for ws in wb.worksheets:
                
                    apply_table_formatting(ws)
    
                wb.save(output)
    
                wb.close()
    
            # -------------------------------------------------
            # SUCCESS LOGGING
            # -------------------------------------------------
    
            self.logger.info(
                "Tracking sheet created: %s",
                output,
            )
    
            messagebox.showinfo(
                "Success",
                f"Tracking sheet created:\n{output}",
            )
    
        except Exception as exc:
        
            self.logger.exception(
                "Update Tracking Sheet failed"
            )
    
            messagebox.showerror(
                "Error",
                str(exc),
            )
    
        finally:
            try:
                mem_ctx.__exit__(None, None, None)
            except Exception:
                pass
            release_large_objects(
                locals(),
                ["raw_df", "master_df", "new_df", "old_df", "unique_df", "wb", "output"],
            )
        
            self.run_btn.configure(
                state="normal"
            )
    
