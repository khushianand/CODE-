"""Window 2: project selection step."""

from tkinter import ttk


class Window2ProjectSelection(ttk.Frame):
    """Stores selected project into shared app state."""
    def __init__(self, master, state, on_prev, on_next):
        super().__init__(master)
        self.state = state
        self.on_prev = on_prev
        self.on_next = on_next
        self.project = state.get("selected_project", "Select Project")
        self._build()

    def _build(self):
        ttk.Label(self, text="Select Project", font=("Segoe UI", 12, "bold")).pack(pady=10)
        self.combo = ttk.Combobox(self, values=["PTA", "3UK", "Airtel", "Antina", "AT&T Mexico", "AT&T US", "Fast-Web", "One NZ"], state="readonly")
        self.combo.set(self.project)
        self.combo.pack(pady=10)
        btns = ttk.Frame(self)
        btns.pack(pady=20)
        ttk.Button(btns, text="Previous", command=self.on_prev).pack(side="left", padx=8)
        ttk.Button(btns, text="Next", command=self._next).pack(side="left", padx=8)

    def _next(self):
        self.state["selected_project"] = self.combo.get()
        self.on_next()
