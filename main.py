"""Application entry point and window-router for the Tkinter UI flow."""

import tkinter as tk
from tkinter import ttk

from gui.window1_entry import Window1Entry
from gui.window2_project_selection import Window2ProjectSelection
from gui.window3_scanner_selection import Window3ScannerSelection
from gui.window4_main import Window4Main
from utils.logger import get_logger


class App(tk.Tk):
    """Root app that stores shared state and navigates between wizard windows."""
    def __init__(self):
        """Explain workflow and purpose of `__init__` in this module."""
        super().__init__()
        self.title("Vulnerability Management Excel Automation")
        self.geometry("1200x780")
        self.state_data = {
            "entry_mode": "VNF",
            "selected_project": "Select Project",
            "selected_scanner": "Nessus",
        }
        self.logger = get_logger()
        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)
        self.current = None
        self.show_window1()

    def _swap(self, frame):
        """Explain workflow and purpose of `_swap` in this module."""
        if self.current:
            self.current.destroy()
        self.current = frame
        self.current.pack(fill="both", expand=True)

    def show_window1(self):
        """Explain workflow and purpose of `show_window1` in this module."""
        self._swap(Window1Entry(self.container, self.state_data, self.show_window2))

    def show_window2(self):
        """Explain workflow and purpose of `show_window2` in this module."""
        self._swap(Window2ProjectSelection(self.container, self.state_data, self.show_window1, self.show_window3))

    def show_window3(self):
        """Explain workflow and purpose of `show_window3` in this module."""
        self._swap(Window3ScannerSelection(self.container, self.state_data, self.show_window2, self.show_window4))

    def show_window4(self):
        """Explain workflow and purpose of `show_window4` in this module."""
        frame = Window4Main(self.container, self.state_data, self.logger)
        self._swap(frame)
        frame.attach_log_handler()


if __name__ == "__main__":
    App().mainloop()
