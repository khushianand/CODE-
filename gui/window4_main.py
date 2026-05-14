"""Window 4: tabbed working area + live log console."""

import tkinter as tk
from tkinter import scrolledtext, ttk

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
        paned = ttk.Panedwindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        upper = ttk.Frame(paned)
        lower = ttk.Frame(paned)
        paned.add(upper, weight=4)
        paned.add(lower, weight=1)

        nb = ttk.Notebook(upper)
        nb.pack(fill="both", expand=True)
        nb.add(MakeNewReportTab(nb, self.state, self.logger), text="Make New Report")
        nb.add(GenerateTrackingTab(nb, self.state, self.logger), text="Update Tracking Sheet")

        ttk.Label(lower, text="Logs (drag separator to resize)").pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(lower, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True)

    def attach_log_handler(self):
        from utils.logger import UILogHandler

        handler = UILogHandler(self.append_log)
        handler.setFormatter(__import__("logging").Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))
        self.logger.addHandler(handler)
        self._attach_stdio_redirects()

    def _attach_stdio_redirects(self):
        import io
        import sys

        class _StreamToLog(io.TextIOBase):
            def __init__(self, callback):
                self.callback = callback

            def write(self, s):
                text = str(s).rstrip()
                if text:
                    self.callback(text)
                return len(s)

        sys.stdout = _StreamToLog(self.append_log)
        sys.stderr = _StreamToLog(self.append_log)

    def append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")
