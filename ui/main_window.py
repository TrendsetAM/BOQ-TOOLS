"""
BOQ Tools Main Window
Comprehensive GUI for Excel file processing and BOQ analysis
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import platform
from typing import List, Dict, Any

# Optional drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# Modern styling
try:
    from ttkthemes import ThemedTk
    THEME_AVAILABLE = True
except ImportError:
    THEME_AVAILABLE = False

# Color coding for confidence
def confidence_color(score):
    if score >= 0.8:
        return '#4CAF50'  # Green
    elif score >= 0.6:
        return '#FFC107'  # Amber
    else:
        return '#F44336'  # Red


def tooltip(widget, text):
    """Simple tooltip for a widget"""
    def on_enter(event):
        widget._tip = tk.Toplevel(widget)
        widget._tip.wm_overrideredirect(True)
        x = event.x_root + 10
        y = event.y_root + 10
        widget._tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(widget._tip, text=text, background="#ffffe0", relief='solid', borderwidth=1, font=("TkDefaultFont", 9))
        label.pack()
    def on_leave(event):
        if hasattr(widget, '_tip'):
            widget._tip.destroy()
            widget._tip = None
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)


class MainWindow:
    def __init__(self, root=None):
        self.root = root or self._create_root()
        self.root.title("BOQ Tools - Excel Processor")
        self.root.geometry("1200x700")
        self.root.minsize(900, 600)
        self._setup_style()
        self._create_menu()
        # Initialize variables before creating widgets
        self.open_files = {}
        self.status_var = tk.StringVar(value="Ready.")
        self.progress_var = tk.DoubleVar(value=0)
        self.status_bar_visible = True
        # Create widgets
        self._create_main_widgets()
        self._setup_drag_and_drop()
        self._bind_shortcuts()
        self._update_status("Welcome to BOQ Tools!")

    def _create_root(self):
        if THEME_AVAILABLE:
            return ThemedTk(theme="arc")
        elif DND_AVAILABLE:
            return TkinterDnD.Tk()
        else:
            return tk.Tk()

    def _setup_style(self):
        style = ttk.Style(self.root)
        if platform.system() == 'Windows':
            style.theme_use('vista')
        elif platform.system() == 'Darwin':
            style.theme_use('aqua')
        else:
            style.theme_use('clam')
        style.configure('TNotebook.Tab', padding=[10, 5])
        style.configure('Treeview', rowheight=28)

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        # File
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_command(label="Export", accelerator="Ctrl+E", command=self.export_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", accelerator="Ctrl+Q", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        # Edit
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self._not_implemented)
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self._not_implemented)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        # View
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Status Bar", command=self._toggle_status_bar)
        menubar.add_cascade(label="View", menu=view_menu)
        # Tools
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Settings", command=self._not_implemented)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        self.root.config(menu=menubar)

    def _create_main_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        # Top: Drag-and-drop zone
        self.drop_zone = ttk.Label(main_frame, text="Drop Excel files here or use File > Open", anchor="center", relief="ridge", padding=20)
        self.drop_zone.pack(fill=tk.X, padx=10, pady=8)
        tooltip(self.drop_zone, "Drag and drop .xlsx or .xls files here to open.")
        # Center: Tabbed interface for files
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        # Bottom: Status bar and progress
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w")
        self.status_label.pack(side=tk.LEFT, padx=8)
        self.progress = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100, length=180)
        self.progress.pack(side=tk.RIGHT, padx=8)

    def _setup_drag_and_drop(self):
        if DND_AVAILABLE:
            try:
                if hasattr(self.root, 'drop_target_register'):
                    self.root.drop_target_register(DND_FILES)
                    self.root.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass
        else:
            # Fallback: Make drop zone clickable
            self.drop_zone.bind('<Button-1>', lambda e: self.open_file())
            self.drop_zone.config(text="Click here to open Excel files or use File > Open")

    def _bind_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-e>', lambda e: self.export_file())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<Control-z>', lambda e: self._not_implemented())
        self.root.bind('<Control-y>', lambda e: self._not_implemented())

    def _update_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _toggle_status_bar(self):
        if self.status_bar_visible:
            self.status_label.pack_forget()
            self.progress.pack_forget()
        else:
            self.status_label.pack(side=tk.LEFT, padx=8)
            self.progress.pack(side=tk.RIGHT, padx=8)
        self.status_bar_visible = not self.status_bar_visible

    def _not_implemented(self):
        messagebox.showinfo("Not Implemented", "This feature is not implemented yet.")

    def open_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filenames = filedialog.askopenfilenames(title="Open Excel File", filetypes=filetypes)
        for file in filenames:
            self._open_excel_file(file)

    def _on_drop(self, event):
        if DND_AVAILABLE:
            files = self.root.tk.splitlist(event.data)
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    self._open_excel_file(file)
                else:
                    self._update_status(f"Unsupported file: {file}")

    def _open_excel_file(self, filepath):
        # Placeholder: Replace with actual ExcelProcessor integration
        filename = os.path.basename(filepath)
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=filename)
        self.notebook.select(tab)
        # Sheet selection and grid preview
        self._populate_file_tab(tab, filepath)
        self.open_files[filepath] = tab
        self._update_status(f"Opened: {filename}")

    def _populate_file_tab(self, tab, filepath):
        # Top: Sheet selection and confidence
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(top_frame, text="Sheet:").pack(side=tk.LEFT)
        sheet_combo = ttk.Combobox(top_frame, values=["Sheet1", "Sheet2"], state="readonly", width=18)
        sheet_combo.pack(side=tk.LEFT, padx=5)
        sheet_combo.current(0)
        # Confidence indicator
        conf_score = 0.85  # Placeholder
        conf_label = ttk.Label(top_frame, text=f"Confidence: {int(conf_score*100)}%", background=confidence_color(conf_score), foreground="white", padding=5)
        conf_label.pack(side=tk.LEFT, padx=10)
        tooltip(conf_label, "Overall confidence score for this sheet.")
        # Sheet classification result
        class_label = ttk.Label(top_frame, text="Type: LineItems", padding=5)
        class_label.pack(side=tk.LEFT, padx=10)
        tooltip(class_label, "Sheet classification result.")
        # Manual override
        override_btn = ttk.Button(top_frame, text="Override Mapping", command=self._not_implemented)
        override_btn.pack(side=tk.RIGHT, padx=5)
        tooltip(override_btn, "Manually override column mappings.")
        # Data grid preview
        grid_frame = ttk.Frame(tab)
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        columns = ["Item No.", "Description", "Unit", "Quantity", "Unit Price", "Total Price"]
        tree = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")
        # Placeholder data
        data = [
            ["1.1", "Excavation", "m³", "150.00", "25.00", "3750.00"],
            ["1.2", "Concrete", "m³", "75.00", "120.00", "9000.00"],
            ["2.1", "Brickwork", "m²", "200.00", "45.00", "9000.00"]
        ]
        for row in data:
            tree.insert('', tk.END, values=row)
        tree.pack(fill=tk.BOTH, expand=True)
        tooltip(tree, "Preview of sheet data. Double-click to edit.")
        # Export button
        export_btn = ttk.Button(tab, text="Export Mapping", command=self.export_file)
        export_btn.pack(side=tk.BOTTOM, pady=8)
        tooltip(export_btn, "Export the current mapping to a file.")

    def export_file(self):
        filetypes = [("JSON", "*.json"), ("All files", "*.*")]
        filename = filedialog.asksaveasfilename(title="Export Mapping", defaultextension=".json", filetypes=filetypes)
        if filename:
            # Placeholder: Replace with actual export logic
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("{\n  \"export\": \"placeholder\"\n}")
            self._update_status(f"Exported mapping to {filename}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = MainWindow()
    app.run() 