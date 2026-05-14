"""Window 1: entry mode selection (CNF / VNF)."""

import tkinter as tk
from tkinter import ttk


class Window1Entry(ttk.Frame):
    """Captures entry mode and advances to project selection."""
    def __init__(self, master, state, on_next):
        """Explain workflow and purpose of `__init__` in this module."""
        super().__init__(master)
        self.state = state
        self.on_next = on_next
        self.mode = tk.StringVar(value="VNF")
        self._build()

    def _build(self):
        """Explain workflow and purpose of `_build` in this module."""
        ttk.Label(self, text="Vulnerability Management Automation", font=("Segoe UI", 14, "bold")).pack(pady=15)
        ttk.Radiobutton(self, text="VNF", variable=self.mode, value="VNF").pack(anchor="w", padx=20)
        ttk.Radiobutton(self, text="CNF", variable=self.mode, value="CNF").pack(anchor="w", padx=20)
        ttk.Button(self, text="Next", command=self._next).pack(pady=20)

    def _next(self):
        """Explain workflow and purpose of `_next` in this module."""
        self.state["entry_mode"] = self.mode.get()
        self.on_next()
