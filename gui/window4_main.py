"""Window 4: tabbed working area + live log console."""

import tkinter as tk
from tkinter import scrolledtext, ttk

from tabs.generate_tracking import GenerateTrackingTab
from tabs.make_new_report import MakeNewReportTab


class Window4Main(ttk.Frame):
    """Hosts all processing tabs and pipes logger events to on-screen logs."""
    def __init__(self, master, state, logger):
        """Handle the init step for this module workflow."""
        super().__init__(master)
        self.state = state
        self.logger = logger
        self._build()

    def _build(self):
        """Handle the build step for this module workflow."""
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
        """Handle the attach log handler step for this module workflow."""
        from utils.logger import UILogHandler

        handler = UILogHandler(self.append_log)
        handler.setFormatter(__import__("logging").Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))
        self.logger.addHandler(handler)
        self._attach_stdio_redirects()

    def _attach_stdio_redirects(self):
        """Handle the attach stdio redirects step for this module workflow."""
        import io
        import sys

        class _StreamToLog(io.TextIOBase):
            def __init__(self, callback):
                """Handle the init step for this module workflow."""
                self.callback = callback

            def write(self, s):
                """Handle the write step for this module workflow."""
                text = str(s).rstrip()
                if text:
                    self.callback(text)
                return len(s)

        sys.stdout = _StreamToLog(self.append_log)
        sys.stderr = _StreamToLog(self.append_log)

    def append_log(self, msg: str):
        """Handle the append log step for this module workflow."""
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")
