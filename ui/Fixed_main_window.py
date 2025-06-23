"""
BOQ Tools Main Window
Comprehensive GUI for Excel file processing and BOQ analysis
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import platform
from pathlib import Path
from typing import List, Dict, Any
import logging
import threading
import dataclasses

# Get a logger for this module
logger = logging.getLogger(__name__)

# Optional imports
try:
    from ttkthemes import ThemedTk
    THEME_AVAILABLE = True
except ImportError:
    THEME_AVAILABLE = False

try:
    import tkinterdnd2 as TkinterDnD
    DND_AVAILABLE = True
    DND_FILES = TkinterDnD.DND_FILES
except ImportError:
    DND_AVAILABLE = False

# Import settings dialog
try:
    from ui.settings_dialog import show_settings_dialog
    SETTINGS_AVAILABLE = True
except ImportError:
    SETTINGS_AVAILABLE = False

# Import sheet categorization dialog
try:
    from ui.sheet_categorization_dialog import show_sheet_categorization_dialog
    SHEET_CATEGORIZATION_AVAILABLE = True
except ImportError:
    SHEET_CATEGORIZATION_AVAILABLE = False

# Import row review dialog
try:
    from ui.row_review_dialog import show_row_review_dialog
    ROW_REVIEW_AVAILABLE = True
except ImportError:
    ROW_REVIEW_AVAILABLE = False

# Import preview dialog
try:
    from ui.preview_dialog import show_preview_dialog
    PREVIEW_AVAILABLE = True
except ImportError:
    PREVIEW_AVAILABLE = False

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
        widget._tip = tk.Toplevel()
        widget._tip.wm_overrideredirect(True)
        x = event.x_root + 10
        y = event.y_root + 10
        widget._tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(widget._tip, text=text, background="#ffffe0", relief='solid', borderwidth=1, font=("TkDefaultFont", 9))
        label.pack()
    def on_leave(event):
        if hasattr(widget, '_tip') and widget._tip is not None:
            widget._tip.destroy()
            widget._tip = None
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)


class MainWindow:
    def __init__(self, controller, root=None):
        self.controller = controller
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
        self.file_mapping = None  # To store the full file mapping object
        self.sheet_treeviews = {} # To store treeview widgets for each sheet
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
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=lambda: self._update_status('Undo is not implemented.'))
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=lambda: self._update_status('Redo is not implemented.'))
        menubar.add_cascade(label="Edit", menu=edit_menu)
        # View
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Status Bar", command=self._toggle_status_bar)
        menubar.add_cascade(label="View", menu=view_menu)
        # Tools
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Settings", command=self.open_settings)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        self.root.config(menu=menubar)

    def _create_main_widgets(self):
        # Configure root window grid
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Main frame using grid
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky=tk.NSEW)
        
        # Configure main_frame grid layout
        main_frame.grid_rowconfigure(0, weight=0)  # Drop zone
        main_frame.grid_rowconfigure(1, weight=1)  # Notebook (expandable)
        main_frame.grid_rowconfigure(2, weight=0)  # Status bar
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Top: Drag-and-drop zone
        self.drop_zone = ttk.Label(main_frame, text="Drop Excel files here or use File > Open", 
                                 anchor="center", relief="ridge", padding=20)
        self.drop_zone.grid(row=0, column=0, sticky=tk.EW, padx=10, pady=8)
        tooltip(self.drop_zone, "Drag and drop .xlsx or .xls files here to open.")
        
        # Center: Tabbed interface for files
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=tk.NSEW, padx=10, pady=5)
        
        # Bottom: Status bar and progress
        self.status_frame = ttk.Frame(main_frame)
        self.status_frame.grid(row=2, column=0, sticky=tk.EW)
        
        # Configure status frame grid
        self.status_frame.grid_columnconfigure(0, weight=1)
        self.status_frame.grid_columnconfigure(1, weight=0)
        
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=0, column=0, sticky=tk.EW, padx=8)
        
        self.progress = ttk.Progressbar(self.status_frame, variable=self.progress_var, maximum=100, length=180)
        self.progress.grid(row=0, column=1, padx=8)

    def _setup_drag_and_drop(self):
        if DND_AVAILABLE:
            try:
                # These attributes are dynamically added by tkinterdnd2
                self.root.drop_target_register(TkinterDnD.DND_FILES)
                self.root.dnd_bind('<<Drop>>', self._on_drop)
            except AttributeError:
                # This can happen if TkinterDnD is available but methods aren't patched correctly (rare)
                logger.warning("TkinterDnD methods not found on root. Drag and drop will not be fully functional.")
                pass
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
        self.root.bind('<Control-z>', lambda e: self._update_status('Undo is not implemented.'))
        self.root.bind('<Control-y>', lambda e: self._update_status('Redo is not implemented.'))

    def _update_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _toggle_status_bar(self):
        if self.status_bar_visible:
            self.status_frame.grid_remove()
        else:
            self.status_frame.grid()
        self.status_bar_visible = not self.status_bar_visible

    def _not_implemented(self):
        # Instead of a dialog, just update the status bar
        self._update_status("This feature is not implemented yet.")

    def open_settings(self):
        if SETTINGS_AVAILABLE:
            show_settings_dialog(self.root)
        else:
            self._not_implemented()

    def export_file(self):
        # Get the currently selected tab
        current_tab = self.notebook.nametowidget(self.notebook.select())
        if hasattr(current_tab, 'file_mapping'):
            # Implement export logic here
            self._update_status("Export functionality to be implemented.")
        else:
            self._update_status("No data to export.")

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
        """Handle the file processing workflow"""
        if not filepath:
            return

        # Clear previous results, including treeview references
        self.sheet_treeviews.clear()
        if self.notebook:
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)

        self._update_status(f"Processing {os.path.basename(filepath)}, please wait...")
        self.progress_var.set(0)

        # Create a new tab for the file
        filename = os.path.basename(filepath)
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=filename)
        self.notebook.select(tab)

        # Configure grid for the tab frame
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        loading_label = ttk.Label(tab, text="Analyzing file...")
        loading_label.grid(row=0, column=0, pady=40, padx=100)
        self.root.update_idletasks()

        def process_in_thread():
            """Runs the file processing in a background thread."""
            try:
                # Step 1: Load the file and get visible sheets
                from core.file_processor import ExcelProcessor
                processor = ExcelProcessor()
                processor.load_file(Path(filepath))
                visible_sheets = processor.get_visible_sheets()
                if not visible_sheets:
                    self.root.after(0, self._on_processing_error, tab, filename, loading_label)
                    return
                
                # Step 2: Ask user to categorize sheets
                def ask_categorization():
                    if SHEET_CATEGORIZATION_AVAILABLE:
                        categories = show_sheet_categorization_dialog(self.root, visible_sheets)
                        if not categories:
                            # User cancelled
                            self._update_status("Sheet categorization cancelled.")
                            loading_label.destroy()
                            return
                        boq_sheets = [sheet for sheet, cat in categories.items() if cat == "BOQ"]
                        if not boq_sheets:
                            self._update_status("No sheets marked as BOQ. Processing aborted.")
                            loading_label.destroy()
                            return
                    else:
                        # Fallback: treat all sheets as BOQ
                        boq_sheets = visible_sheets
                        categories = {sheet: "BOQ" for sheet in visible_sheets}
                    
                    # Step 3: Process only BOQ sheets
                    file_mapping = self.controller.process_file(
                        Path(filepath),
                        progress_callback=lambda p, m: self.root.after(0, self.update_progress, p, m),
                        sheet_filter=boq_sheets,
                        sheet_types=categories
                    )
                    # After processing, show the main window with BOQ sheets for column mapping
                    self.file_mapping = file_mapping
                    self.root.after(0, self._on_processing_complete, tab, filepath, self.file_mapping, loading_label)
                
                self.root.after(0, ask_categorization)
            except Exception as e:
                logger.error(f"Failed to process file {filepath}: {e}", exc_info=True)
                self.root.after(0, self._on_processing_error, tab, filename, loading_label)

        threading.Thread(target=process_in_thread, daemon=True).start()

    def update_progress(self, percentage, message):
        """Thread-safe method to update the progress bar and status label."""
        self.progress_var.set(percentage)
        self._update_status(message)

    def _on_processing_complete(self, tab, filepath, file_mapping, loading_widget):
        """Callback for when file processing succeeds. Runs in the main UI thread."""
        loading_widget.destroy()
        # Store the real mapping data with the tab widget itself for later access (like exporting)
        setattr(tab, 'file_mapping', file_mapping)
        self._populate_file_tab(tab, file_mapping)
        self._update_status(f"Successfully processed: {os.path.basename(filepath)}")
        self.progress_var.set(100)
    
    def _on_processing_error(self, tab, filename, loading_widget):
        """Callback for when file processing fails. Runs in the main UI thread."""
        loading_widget.destroy()
        # Use grid for consistency
        error_label = ttk.Label(tab, text=f"Failed to process {filename}.\nSee logs for details.", foreground="red")
        error_label.grid(row=0, column=0, pady=40, padx=100)
        self._update_status(f"Error processing {filename}")
        self.progress_var.set(0)

    def _populate_file_tab(self, tab, file_mapping):
        """Populates a tab with the processed data from a file mapping."""
        # Debug: print all sheet names and their types
        print('DEBUG: Sheets in file_mapping:')
        for s in file_mapping.sheets:
            print(f'  {s.sheet_name} (sheet_type={getattr(s, "sheet_type", None)})')
        
        # Clear any existing widgets (like loading/error labels)
        for widget in tab.winfo_children():
            widget.destroy()
        
        # Create main container frame
        tab_frame = ttk.Frame(tab)
        tab_frame.grid(row=0, column=0, sticky=tk.NSEW)
        
        # Configure tab_frame's grid layout
        tab_frame.grid_rowconfigure(0, weight=0)  # For export_frame
        tab_frame.grid_rowconfigure(1, weight=0)  # For global_summary
        tab_frame.grid_rowconfigure(2, weight=1)  # For sheet_notebook (main content, expands vertically)
        tab_frame.grid_rowconfigure(3, weight=0)  # For confirm_col_btn
        tab_frame.grid_rowconfigure(4, weight=1)  # For row_review_container (expands vertically)
        tab_frame.grid_columnconfigure(0, weight=1)  # Only one column, expands horizontally

        # Add export button at the top
        export_frame = ttk.Frame(tab_frame)
        export_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        export_frame.grid_columnconfigure(0, weight=1)  # Allow expansion
        
        export_btn = ttk.Button(export_frame, text="Export Data", command=self.export_file)
        export_btn.grid(row=0, column=1, padx=5)

        # Add global summary
        global_summary = ttk.LabelFrame(tab_frame, text="File Summary")
        global_summary.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
        
        global_text = f"""Total Sheets: {len(file_mapping.sheets)}
Global Confidence: {file_mapping.global_confidence:.1%}
Export Ready: {'Yes' if file_mapping.export_ready else 'No'}
Processing Status: {file_mapping.processing_summary.successful_sheets} successful, {file_mapping.processing_summary.partial_sheets} partial"""
        
        summary_label = ttk.Label(global_summary, text=global_text)
        summary_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)

        # Create sheet notebook for individual sheet tabs
        sheet_notebook = ttk.Notebook(tab_frame)
        sheet_notebook.grid(row=2, column=0, sticky=tk.NSEW, padx=5, pady=5)
        
        # Populate each sheet as a tab in the sheet_notebook
        for sheet in file_mapping.sheets:
            sheet_frame = ttk.Frame(sheet_notebook)
            sheet_notebook.add(sheet_frame, text=sheet.sheet_name)
            self._populate_sheet_tab(sheet_frame, sheet)
        
        # Add confirmation button for column mappings
        confirm_frame = ttk.Frame(tab_frame)
        confirm_frame.grid(row=3, column=0, sticky=tk.EW, padx=5, pady=5)
        confirm_frame.grid_columnconfigure(0, weight=1)
        
        confirm_btn = ttk.Button(confirm_frame, text="Apply These Mappings to All Other Sheets", 
                               command=lambda: self._apply_mappings_to_all_sheets(file_mapping))
        confirm_btn.grid(row=0, column=1, padx=5)
        
        # Row review container (placeholder for now)
        review_frame = ttk.LabelFrame(tab_frame, text="Row Review")
        review_frame.grid(row=4, column=0, sticky=tk.NSEW, padx=5, pady=5)
        review_frame.grid_rowconfigure(0, weight=1)
        review_frame.grid_columnconfigure(0, weight=1)
        
        review_label = ttk.Label(review_frame, text="Row review functionality will be implemented here.")
        review_label.grid(row=0, column=0, padx=10, pady=10)

    def _populate_sheet_tab(self, sheet_frame, sheet):
        """Populate an individual sheet tab with its data and column mappings."""
        # Configure sheet frame grid
        sheet_frame.grid_rowconfigure(0, weight=0)  # Summary
        sheet_frame.grid_rowconfigure(1, weight=1)  # Column mappings table
        sheet_frame.grid_columnconfigure(0, weight=1)
        
        # Sheet summary
        summary_frame = ttk.LabelFrame(sheet_frame, text="Sheet Summary")
        summary_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        
        summary_text = f"""Processing Status: {getattr(sheet, 'processing_status', 'Unknown')}
Confidence: {getattr(sheet, 'confidence', 0):.1%}
Data Rows: {getattr(sheet, 'data_rows', 0)}
Columns: {len(getattr(sheet, 'column_mappings', []))}
Validation Score: {getattr(sheet, 'validation_score', 0):.1%}"""
        
        summary_label = ttk.Label(summary_frame, text=summary_text)
        summary_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        
        # Column mappings table
        mappings_frame = ttk.LabelFrame(sheet_frame, text="Column Mappings (Double-click to edit) - Required columns are highlighted")
        mappings_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=5, pady=5)
        mappings_frame.grid_rowconfigure(0, weight=1)
        mappings_frame.grid_columnconfigure(0, weight=1)
        
        # Create treeview for column mappings
        columns = ("Original Header", "Mapped Type", "Confidence", "Required", "Actions")
        tree = ttk.Treeview(mappings_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(mappings_frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(mappings_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout for treeview and scrollbars
        tree.grid(row=0, column=0, sticky=tk.NSEW)
        v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        h_scrollbar.grid(row=1, column=0, sticky=tk.EW)
        
        # Populate treeview with column mappings
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                confidence = getattr(mapping, 'confidence', 0)
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                required = getattr(mapping, 'required', False)
                original_header = getattr(mapping, 'original_header', 'Unknown')
                
                # Determine background color based on confidence
                tags = []
                if confidence >= 0.8:
                    tags.append('high_confidence')
                elif confidence >= 0.6:
                    tags.append('medium_confidence')
                else:
                    tags.append('low_confidence')
                
                if required:
                    tags.append('required')
                
                tree.insert("", tk.END, values=(
                    original_header,
                    mapped_type,
                    f"{confidence:.1%}",
                    "Yes" if required else "No",
                    "Auto-detected"
                ), tags=tags)
        
        # Configure tags for different confidence levels
        tree.tag_configure('high_confidence', background='#e8f5e8')
        tree.tag_configure('medium_confidence', background='#fff3cd')
        tree.tag_configure('low_confidence', background='#f8d7da')
        tree.tag_configure('required', background='#d4edda')
        
        # Store treeview reference for later access
        self.sheet_treeviews[sheet.sheet_name] = tree
        
        # Add legend
        legend_frame = ttk.Frame(mappings_frame)
        legend_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=5)
        
        legend_text = "Legend: Required columns (Description, Quantity, Unit Price, Total Price, Unit, Code) are highlighted and essential for BOQ processing. Actions shows 'Auto-detected' for app decisions and 'Edited' for user changes."
        legend_label = ttk.Label(legend_frame, text=legend_text, wraplength=800, font=("TkDefaultFont", 8))
        legend_label.grid(row=0, column=0, padx=5)

    def _apply_mappings_to_all_sheets(self, file_mapping):
        """Apply the current sheet's column mappings to all other sheets."""
        self._update_status("Applying mappings to all sheets...")
        # This would call the controller to apply mappings
        # For now, just update the status
        self._update_status("Column mappings applied to all sheets.")

    def run(self):
        """Start the main application loop."""
        self.root.mainloop()
