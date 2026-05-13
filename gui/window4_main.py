"""Window 4: tabbed working area + live log console."""

import tkinter as tk
from tkinter import scrolledtext, ttk

from tabs.add_vams_data import AddVamsDataTab
from tabs.generate_tracking import GenerateTrackingTab
from tabs.make_new_report import MakeNewReportTab


class Window4Main(ttk.Frame):
    """Hosts all processing tabs and pipes logger events to on-screen logs."""
    def __init__(self, master, state, logger):
        super().__init__(master)
        self.state = state
        self.logger = logger
        self._build()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        nb.add(MakeNewReportTab(nb, self.state, self.logger), text="Make New Report")
        nb.add(GenerateTrackingTab(nb, self.state, self.logger), text="Update Tracking Sheet")
        nb.add(AddVamsDataTab(nb, self.state, self.logger), text="Add VAMS Data")

        ttk.Label(self, text="Logs").pack(anchor="w", padx=8)
        self.log_text = scrolledtext.ScrolledText(self, height=8, state="disabled")
        self.log_text.pack(fill="x", padx=8, pady=8)

    def attach_log_handler(self):
        from utils.logger import UILogHandler

        handler = UILogHandler(self.append_log)
        handler.setFormatter(__import__("logging").Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))
        self.logger.addHandler(handler)

    def append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")
