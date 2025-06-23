"""
Sheet Categorization Dialog for BOQ Tools
Allows user to categorize each visible sheet as Ignore, BOQ, or Info
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Optional
import platform

class SheetCategorizationDialog:
    def __init__(self, parent, sheet_names: List[str], initial_categories: Optional[Dict[str, str]] = None):
        """
        Args:
            parent: Parent window
            sheet_names: List of visible sheet names
            initial_categories: Optional initial mapping {sheet_name: category}
        """
        self.parent = parent
        self.sheet_names = sheet_names
        self.categories = initial_categories or {name: "BOQ" for name in sheet_names}
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Sheet Categorization")
        self.dialog.geometry("500x400")
        self.dialog.minsize(500, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self._center_dialog()
        self._setup_style()
        self._create_widgets()
        self.dialog.focus_set()

    def _center_dialog(self):
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _setup_style(self):
        style = ttk.Style(self.dialog)
        if platform.system() == 'Windows':
            style.theme_use('vista')
        elif platform.system() == 'Darwin':
            style.theme_use('aqua')
        else:
            style.theme_use('clam')

    def _create_widgets(self):
        self.dialog.grid_rowconfigure(2, weight=1)
        self.dialog.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(self.dialog, text="Categorize Sheets", font=("TkDefaultFont", 14, "bold"))
        title_label.grid(row=0, column=0, sticky="ew", pady=(10, 0), padx=10)

        info_label = ttk.Label(self.dialog, text="Select a category for each visible sheet:")
        info_label.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))

        self.sheet_vars = {}
        # --- Scrollable sheet list ---
        canvas = tk.Canvas(self.dialog, borderwidth=0, highlightthickness=0)
        canvas.grid(row=2, column=0, sticky="nsew", padx=10)
        sheet_frame = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        vsb.grid(row=2, column=1, sticky="ns")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.create_window((0, 0), window=sheet_frame, anchor="nw")
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        sheet_frame.bind("<Configure>", on_frame_configure)
        for sheet in self.sheet_names:
            row = ttk.Frame(sheet_frame)
            row.pack(fill=tk.X, pady=2)
            label = ttk.Label(row, text=sheet, width=30, anchor="w")
            label.pack(side=tk.LEFT)
            var = tk.StringVar(value=self.categories.get(sheet, "BOQ"))
            self.sheet_vars[sheet] = var
            for cat in ["Ignore", "BOQ", "Info"]:
                rb = ttk.Radiobutton(row, text=cat, value=cat, variable=var)
                rb.pack(side=tk.LEFT, padx=5)
        # --- End scrollable sheet list ---
        button_frame = ttk.Frame(self.dialog)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10, padx=10)
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Confirm", command=self._on_confirm).pack(side=tk.RIGHT, padx=5)
        self.dialog.bind('<Return>', lambda e: self._on_confirm())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

    def _on_confirm(self):
        # Collect user choices
        self.result = {sheet: var.get() for sheet, var in self.sheet_vars.items()}
        if not any(cat == "BOQ" for cat in self.result.values()):
            messagebox.showwarning("No BOQ Sheets", "At least one sheet must be categorized as 'BOQ' to proceed.", parent=self.dialog)
            return
        self.dialog.destroy()

    def _on_cancel(self):
        self.result = None
        self.dialog.destroy()

    def show(self) -> Optional[Dict[str, str]]:
        self.dialog.wait_window()
        return self.result

def show_sheet_categorization_dialog(parent, sheet_names: List[str], initial_categories: Optional[Dict[str, str]] = None) -> Optional[Dict[str, str]]:
    dialog = SheetCategorizationDialog(parent, sheet_names, initial_categories)
    return dialog.show() 