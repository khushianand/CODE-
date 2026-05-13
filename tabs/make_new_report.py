"""Tab 1: Create a new report from a raw file (no master comparison)."""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from logic.comparison_logic import aggregate_unique
from logic.excel_writer import write_output
from logic.excel_writer.formatting import (
    apply_table_formatting,
)
from logic.excel_writer.three_uk_qualys import (
    build_3uk_qualys_total_sheet_df,
    build_3uk_qualys_unique_sheet_df,
)


from logic.parser import parse_scan_file
from utils.file_handler import list_excel_sheets, validate_file


class MakeNewReportTab(ttk.Frame):
    """Parses raw scan, builds unique aggregation, writes standard workbook."""
    def __init__(self, master, app_state, logger):
        super().__init__(master)
        self.state = app_state
        self.logger = logger

        self.raw_file = tk.StringVar()
        self.raw_sheet = tk.StringVar()
        self.output_file = tk.StringVar()


        self._build_ui()

    def _build_ui(self):
        ttk.Label(self, text="Raw file").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(self, textvariable=self.raw_file, width=72).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(self, text="Browse", command=self._browse_raw).grid(row=0, column=2, padx=6, pady=6)

        ttk.Label(self, text="Raw sheet").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.raw_sheet_combo = ttk.Combobox(self, textvariable=self.raw_sheet, state="readonly", width=35)
        self.raw_sheet_combo.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(self, text="Output file").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(self, textvariable=self.output_file, width=72).grid(row=2, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(self, text="Browse", command=self._browse_output).grid(row=2, column=2, padx=6, pady=6)

        self.run_btn = ttk.Button(self, text="Run", command=self.run)
        self.run_btn.grid(row=3, column=1, sticky="e", padx=6, pady=10)
        self.columnconfigure(1, weight=1)

    def _browse_raw(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if not path:
            return
        self.raw_file.set(path)
        sheets = list_excel_sheets(path)
        self.raw_sheet_combo["values"] = sheets
        if sheets:
            self.raw_sheet.set(sheets[0])

    def _browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.output_file.set(path)

    def _validate_inputs(self):
        validate_file(self.raw_file.get())
        if not self.raw_sheet.get():
            raise ValueError("Please select a raw sheet")
        if not self.output_file.get():
            raise ValueError("Please select output file")

    def run(self):
        try:
        
            self.run_btn.configure(
                state="disabled"
            )

            self._validate_inputs()

            # -------------------------------------------------
            # SPECIAL FLOW:
            # 3UK + QUALYS
            # -------------------------------------------------

            if (
            
                self.state["selected_project"]
                .strip()
                .casefold()
                == "3uk"

                and

                self.state["selected_scanner"]
                .strip()
                .casefold()
                == "qualys"
            ):

                # ---------------------------------------------
                # TOTAL VULNERABILITIES DF
                # ---------------------------------------------

                total_df = (
                    build_3uk_qualys_total_sheet_df(
                        self.raw_file.get(),
                        self.raw_sheet.get(),
                    )
                )

                # ---------------------------------------------
                # UNIQUE VULNERABILITIES DF
                # ---------------------------------------------

                unique_df = (
                    build_3uk_qualys_unique_sheet_df(
                        total_df
                    )
                )

                # ---------------------------------------------
                # DASHBOARD SOURCE
                # ---------------------------------------------

                summary_df = total_df

                # ---------------------------------------------
                # WRITE OUTPUT
                # ---------------------------------------------

                output = write_output(
                
                    self.output_file.get(),

                    total_df,

                    total_df.iloc[0:0],

                    unique_df,

                    self.state["selected_project"],

                    self.state["selected_scanner"],

                    include_old_sheet=False,

                    new_sheet_name="Total Vulnerabilities",

                    include_dashboard_sheet=True,
                )

            # -------------------------------------------------
            # GENERIC FLOW
            # -------------------------------------------------

            else:
            
                raw_df = parse_scan_file(
                    self.raw_file.get(),
                    self.raw_sheet.get(),
                    self.state["selected_scanner"],
                    self.state["selected_project"],
                ).df

                unique_df = aggregate_unique(
                    raw_df
                )

                summary_df = raw_df

                output = write_output(
                
                    self.output_file.get(),

                    summary_df,

                    summary_df.iloc[0:0],

                    unique_df,

                    self.state["selected_project"],

                    self.state["selected_scanner"],

                    include_old_sheet=False,

                    new_sheet_name="Total Vulnerabilities",

                    include_dashboard_sheet=True,
                )

            # -------------------------------------------------
            # APPLY FORMATTING
            # -------------------------------------------------

            from openpyxl import (
                load_workbook
            )

            wb = load_workbook(output)

            for ws in wb.worksheets:
            
                apply_table_formatting(ws)

            wb.save(output)

            wb.close()

            # -------------------------------------------------

            self.logger.info(
                "Make New Report generated: %s",
                output,
            )

            messagebox.showinfo(
                "Success",
                f"Report generated:\n{output}",
            )

        except Exception as exc:
        
            self.logger.exception(
                "Make New Report failed"
            )

            messagebox.showerror(
                "Error",
                str(exc),
            )

        finally:
        
            self.run_btn.configure(
                state="normal"
            )