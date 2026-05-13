"""Tab 3: merge VAMS data into an existing generated workbook."""

import tkinter as tk

from tkinter import (
    filedialog,
    messagebox,
    ttk,
)

from openpyxl import (
    load_workbook,
)

from logic.fast_vams_enrichment import (
    FastVamsEnrichmentEngine,
)

from logic.excel_writer import (
    read_sheet_as_df,
    write_summary_sheet,
)

from logic.parser import (
    parse_scan_file,
)

from logic.vams_enrichment import (
    VAMS_COLUMNS,
)

from logic.excel_writer.three_uk_qualys import (
    build_3uk_vams_matching_df,
)

from logic.excel_writer.formatting import (
    apply_table_formatting,
)

from utils.file_handler import (
    list_excel_sheets,
    validate_file,
)


TOTAL_SHEET_CANDIDATES = (

    "Total Vulnerabilities",

    "Total vulnerabilities",

    "Total Data",

    "New Data",

    "Total Vulnerability",
)

UNIQUE_SHEET_CANDIDATES = (

    "Unique Vulnerabilities",

    "Unique Data",
)


class AddVamsDataTab(ttk.Frame):

    def __init__(
        self,
        master,
        app_state,
        logger,
    ):

        super().__init__(master)

        self.state = app_state

        self.logger = logger

        self.output_file = tk.StringVar()

        self.raw_file = tk.StringVar()

        self.raw_sheet = tk.StringVar()

        self._build()


    def _build(self):

        ttk.Label(
            self,
            text="Output file",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=6,
            pady=6,
        )

        ttk.Entry(
            self,
            textvariable=self.output_file,
            width=72,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=6,
            pady=6,
        )

        ttk.Button(
            self,
            text="Browse",
            command=self._browse_output,
        ).grid(
            row=0,
            column=2,
            padx=6,
            pady=6,
        )

        ttk.Label(
            self,
            text="Raw VAMS file",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=6,
            pady=6,
        )

        ttk.Entry(
            self,
            textvariable=self.raw_file,
            width=72,
        ).grid(
            row=2,
            column=1,
            sticky="ew",
            padx=6,
            pady=6,
        )

        ttk.Button(
            self,
            text="Browse",
            command=self._browse_raw,
        ).grid(
            row=2,
            column=2,
            padx=6,
            pady=6,
        )

        ttk.Label(
            self,
            text="Raw VAMS sheet",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            padx=6,
            pady=6,
        )

        self.raw_sheet_combo = ttk.Combobox(
            self,
            textvariable=self.raw_sheet,
            state="readonly",
            width=35,
        )

        self.raw_sheet_combo.grid(
            row=3,
            column=1,
            sticky="w",
            padx=6,
            pady=6,
        )

        self.run_btn = ttk.Button(
            self,
            text="Run",
            command=self.run,
        )

        self.run_btn.grid(
            row=3,
            column=1,
            sticky="e",
            padx=6,
            pady=10,
        )

        self.columnconfigure(
            1,
            weight=1,
        )

    def _browse_output(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        self.output_file.set(path)    


    def _browse_raw(self):
        path = filedialog.askopenfilename(
            filetypes=[
                (
                    "Excel",
                    "*.xlsx *.xls *.xlsm",
                )
            ]
        )
        if not path:
            return
        self.raw_file.set(path)
        sheets = list_excel_sheets(path)
        self.raw_sheet_combo["values"] = sheets
        self.raw_sheet.set(
            sheets[0]
        )    
    def _validate_inputs(self):
        validate_file(
            self.output_file.get()
        )
        validate_file(
            self.raw_file.get()
        )
        if not self.raw_sheet.get():
            raise ValueError(
                "Please select raw VAMS sheet"
            )     
        
    def _read_generated_sheet(
        self,
        path: str,
        *sheet_names: str,
    ):
        for sheet_name in sheet_names:
            try:
                return read_sheet_as_df(
                    path,
                    sheet_name,
                )
            except ValueError:
                continue
        return None
    
    def _normalize_match_value(
        self,
        value,
    ):

        if value is None:
            return ""

        value = str(value)

        value = value.strip()

        value = value.lower()

        value = value.replace(
            "\n",
            " ",
        )

        value = value.replace(
            "\r",
            " ",
        )

        value = " ".join(
            value.split()
        )

        return value


    def _write_vams_columns_only(
        self,
        path: str,
        sheet_name: str,
        enriched_df,
    ):
        self.logger.info(
            "Opening workbook for VAMS writeback: %s",
            sheet_name,
        )
        wb = load_workbook(path)
        ws = wb[sheet_name]
        header_row = 2
        data_start_row = 3
        # ---------------------------------------------------
        # READ HEADERS
        # ---------------------------------------------------
        headers = {}
        for col_idx in range(
            1,
            ws.max_column + 1,
        ):
            header_value = ws.cell(
                row=header_row,
                column=col_idx,
            ).value
            if header_value is None:
                continue
            headers[
                str(header_value).strip()
            ] = col_idx
        # ---------------------------------------------------
        # REQUIRED MATCHING COLUMNS
        # ---------------------------------------------------
        required_columns = [
            "Name",
            "Host/Image",
            "Port",
            "CVE",
            "Scanner ID",
        ]
        for col in required_columns:
            if col not in headers:
                self.logger.error(
                    "Missing required column in Excel sheet: %s",
                    col,
                )
                return
        # ---------------------------------------------------
        # BUILD EXCEL LOOKUP
        # ---------------------------------------------------
        excel_lookup = {}
        for excel_row in range(
            data_start_row,
            ws.max_row + 1,
        ):
            try:
                name = self._normalize_match_value(

                    ws.cell(
                        excel_row,
                        headers["Name"],
                    ).value
                )
                # ---------------------------------------------
                # HOST COLUMN FALLBACKS
                # ---------------------------------------------
                host_column = None
                for possible in [
                    "Host / Image",
                    "Host",
                    "IP",
                ]:
                    if possible in headers:
                        host_column = (
                            headers[possible]
                        )
                        break
                host = ""
                if host_column is not None:
                    host = str(
                        ws.cell(
                            excel_row,
                            host_column,
                        ).value or ""
                    ).strip().lower()
                port = str(
                    ws.cell(
                        excel_row,
                        headers["Port"],
                    ).value or ""
                ).strip().lower()
                cve = str(
                    ws.cell(
                        excel_row,
                        headers["CVE"],
                    ).value or ""
                ).strip().lower()
                key = (
                    f"{name}|"
                    f"{host}|"
                    f"{port}|"
                    f"{cve}"
                )
                excel_lookup[key] = (
                    excel_row
                )
            except Exception:
                continue
        # ---------------------------------------------------
        # NORMALIZE ENRICHED DF
        # ---------------------------------------------------
        df = (
            enriched_df
            .fillna("")
            .astype(str)
            .copy()
        )
        updated = 0
        # ---------------------------------------------------
        # WRITE VAMS VALUES
        # ---------------------------------------------------
        for _, row in df.iterrows():
            try:
                name = (
                    self._normalize_match_value(
                        row.get(
                            "Name",
                            "",
                        )
                    )
                )

                host = (
                    self._normalize_match_value(

                        row.get(

                            "Host / Image",

                            row.get(

                                "Host",

                                row.get(
                                    "IP",
                                    "",
                                ),
                            ),
                        )
                    )
                )

                port = (
                    self._normalize_match_value(
                        row.get(
                            "Port",
                            "",
                        )
                    )
                )

                cve = (
                    self._normalize_match_value(
                        row.get(
                            "CVE",
                            "",
                        )
                    )
                )

                key = (
                    f"{name}|"
                    f"{host}|"
                    f"{port}|"
                    f"{cve}"
                )

                excel_row = (
                    excel_lookup.get(key)
                )
                if not excel_row:
                    continue
                row_updated = False
                # ---------------------------------------------
                # WRITE ONLY VAMS COLUMNS
                # ---------------------------------------------
                for col_name in (
                    VAMS_COLUMNS
                ):
                    col_idx = headers.get(
                        col_name
                    )
                    if col_idx is None:
                        continue
                    value = str(
                        row.get(
                            col_name,
                            "",
                        )
                    ).strip()
                    if not value:
                        continue
                    ws.cell(
                        row=excel_row,
                        column=col_idx,
                        value=value,
                    )
                    row_updated = True
                if row_updated:
                    updated += 1
            except Exception:
                continue
        # ---------------------------------------------------
        # FORMAT SHEET
        # ---------------------------------------------------
        apply_table_formatting(ws)
        # ---------------------------------------------------
        # SAVE WORKBOOK
        # ---------------------------------------------------
        self.logger.info(
            "Saving workbook..."
        )
        wb.save(path)
        wb.close()
        self.logger.info(
            "Updated %s rows in %s",
            updated,
            sheet_name,
        )
    def _refresh_dashboard_charts(
        self,
        path: str,
        vams_chart_df,
    ):
        # -------------------------------------------------
        # READ SHEETS
        # -------------------------------------------------
        total_df = self._read_generated_sheet(
            path,
            *TOTAL_SHEET_CANDIDATES,
        )
        unique_df = self._read_generated_sheet(
            path,
            *UNIQUE_SHEET_CANDIDATES,
        )
        if (
            total_df is None
            or unique_df is None
        ):
            self.logger.warning(
                "Dashboard refresh skipped because required sheets were not found"
            )
            return
        # -------------------------------------------------
        # LOAD WORKBOOK
        # -------------------------------------------------
        wb = load_workbook(path)
        # -------------------------------------------------
        # GET DASHBOARD SHEET
        # -------------------------------------------------
        if "Dashboard" in wb.sheetnames:
            ws = wb["Dashboard"]
        elif "Summary Data" in wb.sheetnames:
            ws = wb["Summary Data"]
            ws.title = "Dashboard"
        else:
            ws = wb.create_sheet(
                "Dashboard",
                0,
            )
        # -------------------------------------------------
        # CLEAR EXISTING DASHBOARD
        # -------------------------------------------------
        ws.delete_rows(
            1,
            ws.max_row,
        )
        # -------------------------------------------------
        # REBUILD DASHBOARD
        # -------------------------------------------------
        write_summary_sheet(
            ws,
            total_df,
            total_df.iloc[0:0],
            unique_df,
            self.state["selected_project"],
            self.state["selected_scanner"],
            include_old_summary=False,
            include_vams_charts=True,
            vams_chart_df=vams_chart_df,
        )
        # -------------------------------------------------
        # SAVE
        # -------------------------------------------------
        wb.save(path)
        wb.close()
        self.logger.info(
            "Dashboard refreshed successfully"
        )

    def run(self):

        try:

            self.run_btn.configure(
                state="disabled"
            )

            self._validate_inputs()

            self.logger.info(
                "Loading workbook sheets..."
            )

            total_df = (
                self._read_generated_sheet(
                    self.output_file.get(),
                    *TOTAL_SHEET_CANDIDATES,
                )
            )

            unique_df = (
                self._read_generated_sheet(
                    self.output_file.get(),
                    *UNIQUE_SHEET_CANDIDATES,
                )
            )

            if total_df is None:

                raise ValueError(
                    "Could not find Total Vulnerabilities sheet"
                )

            if unique_df is None:

                raise ValueError(
                    "Could not find Unique Vulnerabilities sheet"
                )

            # -------------------------------------------------
            # PARSE INCOMING VAMS FILE
            # -------------------------------------------------

            self.logger.info(
                "Parsing incoming VAMS file..."
            )

            incoming_vams = parse_scan_file(

                self.raw_file.get(),

                self.raw_sheet.get(),

                self.state["selected_scanner"],

                self.state["selected_project"],

                require_rows_after_filter=False,

                apply_severity_filter=False,

            ).df

            if incoming_vams.empty:

                messagebox.showwarning(
                    "Warning",
                    "No rows found in VAMS file",
                )

                return

            # -------------------------------------------------
            # INITIALIZE FAST ENGINE
            # -------------------------------------------------

            self.logger.info(
                "Initializing enrichment engine..."
            )

            engine = (
                FastVamsEnrichmentEngine()
            )

            engine.build_vams_lookup(
                incoming_vams
            )

            # -------------------------------------------------
            # SPECIAL FLOW:
            # 3UK + QUALYS
            # -------------------------------------------------

            is_3uk_qualys = (

                self.state["selected_project"]
                .strip()
                .casefold()
                == "3uk"

                and

                self.state["selected_scanner"]
                .strip()
                .casefold()
                == "qualys"
            )

            # -------------------------------------------------
            # ENRICH UNIQUE DF
            # -------------------------------------------------

            self.logger.info(
                "Enriching Unique Vulnerabilities..."
            )

            enriched_unique_df = (
                engine.enrich(
                    unique_df
                )
            )

            # -------------------------------------------------
            # AVAILABLE SHEETS
            # -------------------------------------------------

            available_sheets = (
                list_excel_sheets(
                    self.output_file.get()
                )
            )

            target_unique_sheet = next(

                (
                    sheet_name

                    for sheet_name
                    in UNIQUE_SHEET_CANDIDATES

                    if sheet_name
                    in available_sheets
                ),

                "Unique Vulnerabilities",
            )

            target_total_sheet = next(

                (
                    sheet_name

                    for sheet_name
                    in TOTAL_SHEET_CANDIDATES

                    if sheet_name
                    in available_sheets
                ),

                "Total Vulnerabilities",
            )

            # -------------------------------------------------
            # GENERIC FLOW
            # ALL SCANNERS EXCEPT:
            # 3UK + QUALYS
            # -------------------------------------------------

            if not is_3uk_qualys:

                self.logger.info(
                    "Generic flow detected."
                )

                # ---------------------------------------------
                # ENRICH TOTAL
                # ---------------------------------------------

                self.logger.info(
                    "Enriching Total Vulnerabilities..."
                )

                enriched_total_df = (
                    engine.enrich(
                        total_df
                    )
                )

                # ---------------------------------------------
                # WRITE TOTAL
                # ---------------------------------------------

                self.logger.info(
                    "Writing VAMS data into Total Vulnerabilities..."
                )

                self._write_vams_columns_only(

                    self.output_file.get(),

                    target_total_sheet,

                    enriched_total_df,
                )

                # ---------------------------------------------
                # WRITE UNIQUE
                # ---------------------------------------------

                self.logger.info(
                    "Writing VAMS data into Unique Vulnerabilities..."
                )

                self._write_vams_columns_only(

                    self.output_file.get(),

                    target_unique_sheet,

                    enriched_unique_df,
                )

            # -------------------------------------------------
            # SPECIAL FLOW:
            # 3UK + QUALYS
            # -------------------------------------------------

            else:

                self.logger.info(
                    "3UK + Qualys detected."
                )

                self.logger.info(
                    "Skipping Total Vulnerabilities write."
                )

                self.logger.info(
                    "Writing ONLY Unique Vulnerabilities..."
                )

                # ---------------------------------------------
                # ONLY UNIQUE
                # ---------------------------------------------

                self._write_vams_columns_only(

                    self.output_file.get(),

                    target_unique_sheet,

                    enriched_unique_df,
                )
            

            # -------------------------------------------------
            # REFRESH DASHBOARD
            # -------------------------------------------------

            self.logger.info(
                "Refreshing dashboard..."
            )

            self._refresh_dashboard_charts(

                self.output_file.get(),

                enriched_unique_df,
            )

            # -------------------------------------------------
            # COMPLETE
            # -------------------------------------------------

            output = self.output_file.get()

            self.logger.info(
                "VAMS merge completed: %s",
                output,
            )

            messagebox.showinfo(
                "Success",
                "VAMS data merged successfully",
            )

        except Exception as exc:

            self.logger.exception(
                "Add VAMS Data failed"
            )

            messagebox.showerror(
                "Error",
                str(exc),
            )

        finally:

            self.run_btn.configure(
                state="normal"
            )