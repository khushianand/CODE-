"""Window 3: scanner selection step (Nessus / Qualys / Anchor)."""

from tkinter import ttk


class Window3ScannerSelection(ttk.Frame):
    """Stores selected scanner into shared app state."""
    def __init__(self, master, state, on_prev, on_next):
        """Handle the init step for this module workflow."""
        super().__init__(master)
        self.state = state
        self.on_prev = on_prev
        self.on_next = on_next
        self.scanner = state.get("selected_scanner", "Nessus")
        self._build()

    def _build(self):
        """Handle the build step for this module workflow."""
        ttk.Label(self, text="Select Scanner", font=("Segoe UI", 12, "bold")).pack(pady=10)
        self.combo = ttk.Combobox(self, values=["Nessus", "Qualys", "Anchor"], state="readonly")
        self.combo.set(self.scanner)
        self.combo.pack(pady=10)
        btns = ttk.Frame(self)
        btns.pack(pady=20)
        ttk.Button(btns, text="Previous", command=self.on_prev).pack(side="left", padx=8)
        ttk.Button(btns, text="Next", command=self._next).pack(side="left", padx=8)

    def _next(self):
        """Handle the next step for this module workflow."""
        self.state["selected_scanner"] = self.combo.get()
        self.on_next()
