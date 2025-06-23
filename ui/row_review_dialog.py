"""
Row Review Dialog for BOQ Tools
Allows user to review and toggle validity of BOQ rows for each sheet
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional
import platform

class RowReviewDialog:
    def __init__(self, parent, sheet_rows: Dict[str, List[Dict[str, Any]]], required_columns: List[str], initial_valid: Optional[Dict[str, set]] = None):
        """
        Args:
            parent: Parent window
            sheet_rows: {sheet_name: [row_dict, ...]} for each BOQ sheet
            required_columns: List of required column names to display
            initial_valid: Optional {sheet_name: set(row_indices)} of initially valid rows
        """
        self.parent = parent
        self.sheet_rows = sheet_rows
        self.required_columns = required_columns
        self.valid_rows = initial_valid or {sheet: set(range(len(rows))) for sheet, rows in sheet_rows.items()}
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Review BOQ Rows")
        self.dialog.geometry("900x600")
        self.dialog.minsize(700, 400)
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
        style.map('Treeview', background=[('selected', '#B3E5FC')])
        style.configure('validrow', background='#E8F5E9')
        style.configure('invalidrow', background='#FFEBEE')

    def _create_widgets(self):
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        title_label = ttk.Label(main_frame, text="Review BOQ Rows", font=("TkDefaultFont", 14, "bold"))
        title_label.pack(pady=(0, 10))

        info_label = ttk.Label(main_frame, text="Click a row to toggle its validity. Valid rows are highlighted.")
        info_label.pack(anchor=tk.W, pady=(0, 10))

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.treeviews = {}
        for sheet, rows in self.sheet_rows.items():
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=sheet)
            columns = ["#"] + self.required_columns
            tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse", height=18)
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=120 if col != "#" else 40, anchor=tk.W)
            tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            self.treeviews[sheet] = tree
            # Populate rows
            for idx, row in enumerate(rows):
                values = [idx+1] + [row.get(col, "") for col in self.required_columns]
                tag = 'validrow' if idx in self.valid_rows.get(sheet, set()) else 'invalidrow'
                tree.insert('', 'end', iid=str(idx), values=values, tags=(tag,))
            tree.tag_configure('validrow', background='#E8F5E9')
            tree.tag_configure('invalidrow', background='#FFEBEE')
            tree.bind('<Button-1>', lambda e, s=sheet, t=tree: self._on_row_click(e, s, t))

        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Confirm", command=self._on_confirm).pack(side=tk.RIGHT, padx=5)

        self.dialog.bind('<Return>', lambda e: self._on_confirm())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

    def _on_row_click(self, event, sheet, tree):
        region = tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        if idx in self.valid_rows.get(sheet, set()):
            self.valid_rows[sheet].remove(idx)
            tree.item(row_id, tags=('invalidrow',))
        else:
            self.valid_rows[sheet].add(idx)
            tree.item(row_id, tags=('validrow',))

    def _on_confirm(self):
        self.result = {sheet: set(self.valid_rows[sheet]) for sheet in self.sheet_rows}
        self.dialog.destroy()

    def _on_cancel(self):
        self.result = None
        self.dialog.destroy()

    def show(self) -> Optional[Dict[str, set]]:
        self.dialog.wait_window()
        return self.result

def show_row_review_dialog(parent, sheet_rows: Dict[str, List[Dict[str, Any]]], required_columns: List[str], initial_valid: Optional[Dict[str, set]] = None) -> Optional[Dict[str, set]]:
    dialog = RowReviewDialog(parent, sheet_rows, required_columns, initial_valid)
    return dialog.show() 