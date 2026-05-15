"""Window 1: entry mode selection (CNF / VNF)."""

import tkinter as tk
from tkinter import ttk


class Window1Entry(ttk.Frame):
    """Captures entry mode and advances to project selection."""
    def __init__(self, master, state, on_next):
        """Handle the init step for this module workflow."""
        super().__init__(master)
        self.state = state
        self.on_next = on_next
        self.mode = tk.StringVar(value="VNF")
        self._build()

    def _build(self):
        """Handle the build step for this module workflow."""
        ttk.Label(self, text="Vulnerability Management Automation", font=("Segoe UI", 14, "bold")).pack(pady=15)
        ttk.Radiobutton(self, text="VNF", variable=self.mode, value="VNF").pack(anchor="w", padx=20)
        ttk.Radiobutton(self, text="CNF", variable=self.mode, value="CNF").pack(anchor="w", padx=20)
        ttk.Button(self, text="Next", command=self._next).pack(pady=20)

    def _next(self):
        """Handle the next step for this module workflow."""
        self.state["entry_mode"] = self.mode.get()
        self.on_next()
