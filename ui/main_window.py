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
import pandas as pd
from core.row_classifier import RowType
from datetime import datetime
import pickle
import re

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

# Import categorization dialogs
try:
    from ui.categorization_dialog import show_categorization_dialog
    from ui.category_review_dialog import show_category_review_dialog
    from ui.categorization_stats_dialog import show_categorization_stats_dialog
    CATEGORIZATION_AVAILABLE = True
except ImportError:
    CATEGORIZATION_AVAILABLE = False

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


def format_number(value, is_currency=False):
    """Format number with thousands separator and decimal formatting"""
    if not value or value == "":
        return ""
    
    try:
        # Convert to float
        num = float(str(value).replace(',', '').replace(' ', ''))
        
        if is_currency:
            # Format as currency with 2 decimal places
            return f"{num:,.2f}".replace(',', ' ').replace('.', ',')
        else:
            # Format as number with 2 decimal places if needed
            if num == int(num):
                return f"{int(num):,}".replace(',', ' ')
            else:
                return f"{num:,.2f}".replace(',', ' ').replace('.', ',')
    except (ValueError, TypeError):
        return str(value)


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
        self.file_mapping = None
        self.sheet_treeviews = {}
        self.column_mapper = None  # Will be set when processing files
        self.row_validity = {}  # Initialize row validity dictionary
        self.row_review_treeviews = {}  # Initialize row review treeviews dictionary
        # Add robust tab-to-file-mapping
        self.tab_id_to_file_mapping = {}
        # Offer name for summary grid
        self.current_offer_name = None
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
        main_frame.grid_rowconfigure(0, weight=0)  # Buttons zone
        main_frame.grid_rowconfigure(1, weight=1)  # Notebook (expandable)
        main_frame.grid_rowconfigure(2, weight=0)  # Status bar
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Top: Button frame with three buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=0, column=0, sticky=tk.EW, padx=10, pady=8)
        
        # Configure button frame columns to be equal width
        button_frame.grid_columnconfigure(0, weight=1)  # New Analysis
        button_frame.grid_columnconfigure(1, weight=1)  # Load Analysis
        button_frame.grid_columnconfigure(2, weight=1)  # Use Mapping
        
        # Create the three buttons
        new_analysis_btn = ttk.Button(button_frame, text="New Analysis", command=self.open_file)
        new_analysis_btn.grid(row=0, column=0, sticky=tk.EW, padx=5)
        tooltip(new_analysis_btn, "Start a new BOQ analysis by selecting Excel files")
        
        load_analysis_btn = ttk.Button(button_frame, text="Load Analysis", command=self._load_analysis)
        load_analysis_btn.grid(row=0, column=1, sticky=tk.EW, padx=5)
        tooltip(load_analysis_btn, "Load a previously saved analysis")
        
        use_mapping_btn = ttk.Button(button_frame, text="Use Mapping", command=self._use_mapping)
        use_mapping_btn.grid(row=0, column=2, sticky=tk.EW, padx=5)
        tooltip(use_mapping_btn, "Apply a previously saved mapping to new files")
        
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
                # Enable drag and drop if available
                if hasattr(self.root, 'drop_target_register'):
                    self.root.drop_target_register(DND_FILES)
                if hasattr(self.root, 'dnd_bind'):
                    self.root.dnd_bind('<<Drop>>', self._on_drop)
            except AttributeError:
                # This can happen if TkinterDnD is available but methods aren't patched correctly (rare)
                logger.warning("TkinterDnD methods not found on root. Drag and drop will not be fully functional.")
                pass
            except Exception:
                pass
        # No else clause needed - buttons are already set up in _create_main_widgets

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
        # Prompt for offer name/label before opening file dialog
        offer_name = self._prompt_offer_name()
        if offer_name is None:
            self._update_status("File open cancelled (no offer name provided).")
            return
        self.current_offer_name = offer_name
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
                    self.column_mapper = file_mapping.column_mapper if hasattr(file_mapping, 'column_mapper') else None
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
        """Handle processing completion"""
        # Store the file mapping and column mapper
        self.file_mapping = file_mapping
        self.column_mapper = file_mapping.column_mapper if hasattr(file_mapping, 'column_mapper') else None
        
        # Remove loading widget and populate tab
        loading_widget.destroy()
        self._populate_file_tab(tab, file_mapping)
        
        # Update status
        self._update_status(f"Processing complete: {os.path.basename(filepath)}")

    def _on_processing_error(self, tab, filename, loading_widget):
        """Callback for when file processing fails. Runs in the main UI thread."""
        loading_widget.destroy()
        # Use grid for consistency
        error_label = ttk.Label(tab, text=f"Failed to process {filename}.\nSee logs for details.", foreground="red")
        error_label.grid(row=0, column=0, pady=40, padx=100)
        self._update_status(f"Error processing {filename}")
        self.progress_var.set(0)

    def _populate_file_tab(self, tab, file_mapping):
        print("[DEBUG] _populate_file_tab called for tab:", tab)
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
        confirm_btn = ttk.Button(confirm_frame, text="Confirm Column Mappings", command=self._save_all_mappings_for_all_sheets)
        confirm_btn.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        
        # Row Review section will be created only after column mapping is confirmed
        self.row_review_frame = None
        self.row_review_notebook = None
        self.row_review_treeviews = {}
        self.row_validity = {}

        # Ensure file_mapping knows its tab for later lookup
        file_mapping.tab = tab
        # Store mapping from tab ID to file_mapping for robust lookup
        tab_id = str(tab)
        self.tab_id_to_file_mapping[tab_id] = file_mapping

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
        
        # Add the propagate button to the left, above the Treeview in the mappings_frame
        propagate_btn = ttk.Button(mappings_frame, text="Apply These Mappings to All Other Sheets", command=lambda s=sheet: self._apply_mappings_to_all_sheets(s))
        propagate_btn.grid(row=99, column=0, sticky=tk.W, padx=5, pady=(0, 5))
        
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
        
        # Required types
        if self.column_mapper and hasattr(self.column_mapper, 'config'):
            required_types = {col_type.value for col_type in self.column_mapper.config.get_required_columns()}
        else:
            required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        
        # Populate treeview with column mappings
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                confidence = getattr(mapping, 'confidence', 0)
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                required = mapped_type in required_types
                original_header = getattr(mapping, 'original_header', 'Unknown')
                # Determine if this mapping was user-edited
                if hasattr(mapping, 'is_user_edited'):
                    actions = "Edited" if getattr(mapping, 'is_user_edited', False) else "Auto-detected"
                else:
                    # Only show 'Edited' if confidence==1.0 and the Actions field was set to 'Edited' in the edit dialog, it's user-edited
                    actions = "Edited" if (confidence == 1.0 and hasattr(mapping, 'user_edited') and getattr(mapping, 'user_edited', False)) else "Auto-detected"
                tags = []
                if required:
                    tags.append('required')
                tree.insert("", tk.END, values=(
                    original_header,
                    mapped_type,
                    f"{confidence:.1%}",
                    "Yes" if required else "No",
                    actions
                ), tags=tags)
        
        # Only highlight required fields (light blue)
        tree.tag_configure('required', background='#e0f0ff')
        
        # Store treeview reference for later access
        self.sheet_treeviews[sheet.sheet_name] = tree
        
        # Bind double-click to edit
        tree.bind('<Double-1>', lambda e, t=tree, s=sheet: self._edit_column_mapping(t, s))

    def _edit_column_mapping(self, tree, sheet):
        selection = tree.selection()
        if not selection:
            return
        item = tree.item(selection[0])
        values = item['values']
        if not values:
            return
        column_name = values[0]
        # Find the column mapping object
        col_mapping = None
        for cm in getattr(sheet, 'column_mappings', []):
            if getattr(cm, 'original_header', None) == column_name:
                col_mapping = cm
                break
        if not col_mapping:
            return
        # Radio button options for mapped type
        mapped_type_options = [
            "description", "quantity", "unit_price", "total_price", "unit", "code", "ignore"
        ]
        required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        # Dialog to edit mapped type
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Column: {column_name}")
        dialog.geometry("500x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Main content frame
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_label = tk.Label(main_frame, text=f"Column: {column_name}", font=("Arial", 12, "bold"))
        header_label.pack(pady=(0, 10))
        
        # Current mapping info
        current_type = getattr(col_mapping, 'mapped_type', 'unknown')
        current_confidence = getattr(col_mapping, 'confidence', 0)
        info_text = f"Current: {current_type} (Confidence: {current_confidence:.1%})"
        info_label = ttk.Label(main_frame, text=info_text, foreground="gray")
        info_label.pack(pady=(0, 15))
        
        # Selection label
        ttk.Label(main_frame, text="Select new mapped type:").pack(pady=(0, 10))
        
        # Radio buttons for each mapped type
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(pady=5, fill=tk.X)
        type_var = tk.StringVar(value=current_type)
        
        for opt in mapped_type_options:
            radio_btn = ttk.Radiobutton(radio_frame, text=opt, variable=type_var, value=opt)
            radio_btn.pack(anchor=tk.W, padx=10, pady=2)
            
            # Note: ttk.Radiobutton styling is handled by the theme system
        
        # Learning info frame
        learning_frame = ttk.LabelFrame(main_frame, text="Learning Information")
        learning_frame.pack(fill=tk.X, pady=15)
        
        learning_text = "When you save a required type mapping, it will be added to the system's learning database for future use."
        learning_label = ttk.Label(learning_frame, text=learning_text, wraplength=450, justify=tk.LEFT)
        learning_label.pack(padx=10, pady=10)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save():
            new_type = type_var.get()
            # Check for duplicate required type
            if new_type in required_types:
                for cm in getattr(sheet, 'column_mappings', []):
                    if cm is not col_mapping and getattr(cm, 'mapped_type', None) == new_type:
                        other_col_name = getattr(cm, 'original_header', '')
                        proceed = messagebox.askyesno(
                            "Duplicate Required Type",
                            f"The required type '{new_type}' is already assigned to column '{other_col_name}'.\nIf you continue, column '{other_col_name}' will be set to 'unknown'.\nContinue?"
                        )
                        if not proceed:
                            return
                        # Set the other column to unknown and update the treeview
                        cm.mapped_type = "unknown"
                        cm.confidence = 0.0
                        cm.user_edited = True  # Mark as user-edited since user demoted it
                        # Find the row in the treeview for the other column
                        for row_id in tree.get_children():
                            row_vals = tree.item(row_id)['values']
                            if row_vals and row_vals[0] == other_col_name:
                                tree.set(row_id, column="Mapped Type", value="unknown")
                                tree.set(row_id, column="Required", value="No")
                                tree.set(row_id, column="Confidence", value="0.0%")
                                tree.set(row_id, column="Actions", value="Edited")
                                tree.item(row_id, tags=())
                                break
                        break
            # Update the column mapping
            col_mapping.mapped_type = new_type
            col_mapping.confidence = 1.0  # User is always right
            col_mapping.user_edited = True  # Mark as user-edited
            # Learn from user confirmation for required types
            learning_message = ""
            if new_type in required_types and self.column_mapper:
                try:
                    self.column_mapper.update_canonical_mapping(column_name, new_type)
                    learning_message = f"\n\n✓ Learning: '{column_name}' has been added to the system's mapping database."
                    logger.info(f"Learned new mapping: '{column_name}' -> '{new_type}'")
                except Exception as e:
                    learning_message = f"\n\n⚠ Warning: Failed to save mapping to database: {e}"
                    logger.warning(f"Failed to update canonical mapping: {e}")
            # Update the treeview row
            tree.set(selection[0], column="Mapped Type", value=new_type)
            tree.set(selection[0], column="Confidence", value="100.0%")
            # Update the 'Required' field in the treeview row
            required_val = "Yes" if new_type in required_types else "No"
            tree.set(selection[0], column="Required", value=required_val)
            tree.set(selection[0], column="Actions", value="Edited")
            # Update row highlighting for required
            if required_val == "Yes":
                tree.item(selection[0], tags=("required",))
            else:
                tree.item(selection[0], tags=())
            dialog.destroy()
        
        # Buttons
        save_btn = ttk.Button(button_frame, text="Save & Learn", command=save)
        save_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=dialog.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Set focus to save button
        save_btn.focus_set()
        
        # Bind Enter key to save
        dialog.bind('<Return>', lambda e: save())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _apply_mappings_to_all_sheets(self, source_sheet):
        # Propagate user-edited column mappings to all other sheets with the same Original Header (case and whitespace insensitive)
        if not self.file_mapping or not hasattr(self.file_mapping, 'sheets'):
            self._update_status("No file loaded to propagate mappings.")
            return
        user_edited = [cm for cm in getattr(source_sheet, 'column_mappings', []) if getattr(cm, 'confidence', 0) == 1.0]
        if not user_edited:
            self._update_status("No user-edited columns to propagate.")
            return
        count = 0
        required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        affected_sheets = set()
        for edited_cm in user_edited:
            edited_header = getattr(edited_cm, 'original_header', None)
            if edited_header is None:
                continue
            edited_header_key = edited_header.strip().lower()
            for target_sheet in self.file_mapping.sheets:
                if not hasattr(target_sheet, 'column_mappings'):
                    continue
                # Skip the same column object in the source sheet
                if target_sheet is source_sheet:
                    continue
                for target_cm in target_sheet.column_mappings:
                    target_header = getattr(target_cm, 'original_header', None)
                    if target_header is not None and target_header.strip().lower() == edited_header_key:
                        # If required, demote any other column with this type in the target sheet
                        if edited_cm.mapped_type in required_types:
                            for other_cm in target_sheet.column_mappings:
                                if other_cm is not target_cm and other_cm.mapped_type == edited_cm.mapped_type:
                                    other_cm.mapped_type = "unknown"
                                    other_cm.confidence = 0.0
                                    other_cm.user_edited = True  # Mark as user-edited since user demoted it
                        target_cm.mapped_type = edited_cm.mapped_type
                        target_cm.confidence = 1.0
                        target_cm.user_edited = True  # Mark as user-edited (propagated)
                        count += 1
                        affected_sheets.add(getattr(target_sheet, 'sheet_name', None))
        # Always refresh the current sheet as well
        affected_sheets.add(getattr(source_sheet, 'sheet_name', None))
        # Refresh Treeviews for affected sheets
        for sheet_name in affected_sheets:
            tree = self.sheet_treeviews.get(sheet_name)
            if tree:
                # Find the corresponding sheet object
                target_sheet = next((s for s in self.file_mapping.sheets if getattr(s, 'sheet_name', None) == sheet_name), None)
                if target_sheet:
                    # Clear and repopulate the treeview
                    tree.delete(*tree.get_children())
                    for mapping in target_sheet.column_mappings:
                        confidence = getattr(mapping, 'confidence', 0)
                        mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                        required = mapped_type in required_types
                        original_header = getattr(mapping, 'original_header', 'Unknown')
                        # Determine if this mapping was user-edited
                        if hasattr(mapping, 'is_user_edited'):
                            actions = "Edited" if getattr(mapping, 'is_user_edited', False) else "Auto-detected"
                        else:
                            actions = "Edited" if getattr(mapping, 'user_edited', False) else "Auto-detected"
                        tags = []
                        if required:
                            tags.append('required')
                        tree.insert("", tk.END, values=(
                            original_header,
                            mapped_type,
                            f"{confidence:.1%}",
                            "Yes" if required else "No",
                            actions
                        ), tags=tags)
        messagebox.showinfo("Propagation Complete", f"Propagated {count} column mappings to all other sheets.")
        self._update_status(f"Propagated {count} column mappings to all other sheets.")

    def _save_all_mappings_for_all_sheets(self):
        if not self.file_mapping or not hasattr(self.file_mapping, 'sheets'):
            messagebox.showwarning("No File", "No file loaded.")
            return
        
        # Save column mappings first
        total_saved = 0
        total_failed = 0
        total_already = 0
        for sheet in self.file_mapping.sheets:
            saved, failed, already = self._save_all_mappings(sheet, show_dialogs=False)
            total_saved += saved
            total_failed += failed
            total_already += already
        
        # Trigger row mapping with updated column mappings
        self._trigger_row_mapping()

    def _trigger_row_mapping(self):
        """Trigger row mapping with the current column mappings"""
        if not self.file_mapping or not hasattr(self.file_mapping, 'sheets'):
            self._update_status("No file loaded for row mapping.")
            return
        
        try:
            # Import row classifier here to avoid circular imports
            from core.row_classifier import RowClassifier
            from utils.config import ColumnType
            
            row_classifier = RowClassifier()
            
            # Get the original sheet data from the controller
            if hasattr(self, 'controller') and self.controller:
                # Try to get the original sheet data from the controller
                # Find the file key by matching the file mapping
                file_key = None
                for key, file_data in self.controller.current_files.items():
                    if file_data.get('file_mapping') == self.file_mapping:
                        file_key = key
                        break
                
                if file_key and hasattr(self.controller, 'current_files') and file_key in self.controller.current_files:
                    processor_results = self.controller.current_files[file_key].get('processor_results', {})
                    original_sheet_data = processor_results.get('sheet_data', {})
                else:
                    # Fallback: try to reconstruct from current file mapping
                    original_sheet_data = {}
                    for sheet in self.file_mapping.sheets:
                        # This is a fallback - we might not have the original data
                        original_sheet_data[sheet.sheet_name] = []
            else:
                # Fallback when controller is not available
                original_sheet_data = {}
                for sheet in self.file_mapping.sheets:
                    original_sheet_data[sheet.sheet_name] = []
            
            # Process each sheet for row mapping
            updated_sheets = []
            for sheet in self.file_mapping.sheets:
                sheet_data = original_sheet_data.get(sheet.sheet_name, [])
                if not sheet_data:
                    # Skip if we don't have the original data
                    continue
                
                # Convert column mappings to the format expected by row classifier
                # Use the original column mapping from the main processing pipeline
                column_mapping_dict = {}
                
                # Get the original column mapping from the processor results
                if hasattr(self, 'file_mapping') and hasattr(self.file_mapping, 'column_mapper'):
                    # Use the column mapper's original mapping for this sheet
                    sheet_name = sheet.sheet_name
                    if hasattr(self.file_mapping.column_mapper, 'process_sheet_mapping'):
                        # Get the original mapping result for this sheet
                        sheet_data = original_sheet_data.get(sheet_name, [])
                        if sheet_data:
                            mapping_result = self.file_mapping.column_mapper.process_sheet_mapping(sheet_data)
                            # Use the original 0-based column mapping
                            for mapping in mapping_result.mappings:
                                column_mapping_dict[mapping.column_index] = mapping.mapped_type
                else:
                    # Fallback: use the sheet's column mappings with 0-based conversion
                    for col_mapping in sheet.column_mappings:
                        try:
                            # Convert string column type to ColumnType enum
                            col_type = ColumnType(col_mapping.mapped_type)
                            # Fix: Use 0-based indices for row classifier (subtract 1 from column_index)
                            column_mapping_dict[col_mapping.column_index - 1] = col_type
                        except ValueError:
                            # Skip unknown column types
                            continue
                
                # DEBUG: Log the column mapping and first few rows to diagnose indexing
                # Only debug the "Miscellaneous" sheet where Bank Guarantee is located
                if sheet.sheet_name == "Miscellaneous":
                    print(f"[DEBUG] Column mapping for {sheet.sheet_name}: {column_mapping_dict}")
                    if sheet_data:
                        print(f"[DEBUG] First row data: {sheet_data[0] if len(sheet_data) > 0 else 'No data'}")
                        print(f"[DEBUG] Row data length: {len(sheet_data[0]) if sheet_data else 0}")
                        
                        # Show headers to understand column mapping
                        if hasattr(sheet, 'column_mappings'):
                            print(f"[DEBUG] Headers for {sheet.sheet_name}:")
                            for i, col_mapping in enumerate(sheet.column_mappings):
                                header = getattr(col_mapping, 'original_header', f'Column_{i}')
                                mapped_type = getattr(col_mapping, 'mapped_type', 'unknown')
                                confidence = getattr(col_mapping, 'confidence', 0)
                                print(f"[DEBUG] Column {col_mapping.column_index}: '{header}' -> {mapped_type} (confidence: {confidence:.2f})")
                        
                        # Show first 10 rows to find Bank Guarantee
                        for i in range(min(10, len(sheet_data))):
                            if any('Bank Guarantee' in str(cell) for cell in sheet_data[i]):
                                print(f"[DEBUG] FOUND BANK GUARANTEE at row {i}: {sheet_data[i]}")
                            else:
                                print(f"[DEBUG] Row {i}: {sheet_data[i]}")
                
                # Perform row classification
                row_classification_result = row_classifier.classify_rows(sheet_data, column_mapping_dict)
                
                # Update the sheet's row classifications
                sheet.row_classifications = []
                for row_class in row_classification_result.classifications:
                    from core.mapping_generator import RowClassificationInfo
                    row_info = RowClassificationInfo(
                        row_index=row_class.row_index,
                        row_type=row_class.row_type.value,
                        confidence=row_class.confidence,
                        completeness_score=row_class.completeness_score,
                        hierarchical_level=row_class.hierarchical_level,
                        section_title=row_class.section_title,
                        validation_errors=row_class.validation_errors,
                        reasoning=row_class.reasoning
                    )
                    sheet.row_classifications.append(row_info)
                
                # Update sheet statistics
                sheet.row_count = len(sheet_data)
                sheet.overall_confidence = row_classification_result.overall_quality_score
                
                # Update processing status based on row classification quality
                if row_classification_result.overall_quality_score >= 0.8:
                    sheet.processing_status = "SUCCESS"
                elif row_classification_result.overall_quality_score >= 0.6:
                    sheet.processing_status = "PARTIAL"
                else:
                    sheet.processing_status = "NEEDS_REVIEW"
                
                updated_sheets.append(sheet.sheet_name)
            
            # Update the UI to reflect the new row mappings
            self._refresh_sheet_tabs()
            
            # Show success message
            if updated_sheets:
                self._update_status(f"Row mapping completed for {len(updated_sheets)} sheets.")
                
            # After row mapping is complete and data is available:
            # Show the Row Review section with correct data
            self._show_row_review(self.file_mapping, original_sheet_data)
            
        except Exception as e:
            import logging
            logger = logging.getLogger(__name__)
            logger.error(f"Error during row mapping: {e}", exc_info=True)
            messagebox.showerror(
                "Row Mapping Error",
                f"An error occurred during row mapping:\n{str(e)}\n\n"
                f"Please check the logs for more details."
            )
            self._update_status(f"Row mapping failed: {str(e)}")

    def _refresh_sheet_tabs(self):
        """Refresh the sheet tabs to show updated row mapping information"""
        if not self.file_mapping or not hasattr(self.file_mapping, 'sheets'):
            return
        
        # Find the current file tab
        current_tab = None
        for tab in self.notebook.tabs():
            if hasattr(tab, 'file_mapping') and tab.file_mapping == self.file_mapping:
                current_tab = tab
                break
        
        if current_tab:
            # Clear and repopulate the file tab
            for widget in current_tab.winfo_children():
                widget.destroy()
            
            # Repopulate with updated data
            self._populate_file_tab(current_tab, self.file_mapping)

    def _save_all_mappings(self, sheet, show_dialogs=True):
        """Save all required field mappings from the given sheet to the JSON file if not already present. Returns (saved, failed, already_present)."""
        if not self.column_mapper:
            if show_dialogs:
                messagebox.showwarning("No Column Mapper", "Column mapper not available. Please reload the file.")
            return 0, 0, 0
        required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        saved_count = 0
        failed_count = 0
        already_present_count = 0
        current_mappings = self.column_mapper.get_canonical_mappings()
        for mapping in getattr(sheet, 'column_mappings', []):
            mapped_type = getattr(mapping, 'mapped_type', '')
            if mapped_type in required_types:
                original_header = getattr(mapping, 'original_header', '')
                confidence = getattr(mapping, 'confidence', 0)
                if original_header and mapped_type:
                    is_already_present = False
                    if mapped_type in current_mappings:
                        normalized_header = original_header.strip()
                        for existing_header in current_mappings[mapped_type]:
                            if existing_header.strip() == normalized_header:
                                is_already_present = True
                                break
                    if is_already_present:
                        already_present_count += 1
                        logger.debug(f"Mapping already present: '{original_header}' -> '{mapped_type}'")
                    else:
                        try:
                            self.column_mapper.update_canonical_mapping(original_header, mapped_type)
                            saved_count += 1
                            action_type = "user-edited" if confidence == 1.0 else "auto-detected"
                            logger.info(f"Saved {action_type} mapping: '{original_header}' -> '{mapped_type}' (confidence: {confidence:.1%})")
                        except Exception as e:
                            failed_count += 1
                            logger.warning(f"Failed to save mapping '{original_header}' -> '{mapped_type}': {e}")
        if show_dialogs:
            if saved_count > 0:
                messagebox.showinfo(
                    "Mappings Saved", 
                    f"Successfully saved {saved_count} new mappings to the database.\n"
                    f"These mappings will be used for future file processing.\n\n"
                    f"Already present: {already_present_count} mappings"
                )
                self._update_status(f"Saved {saved_count} new mappings to database.")
            else:
                if already_present_count > 0:
                    messagebox.showinfo(
                        "No New Mappings to Save", 
                        f"All {already_present_count} required field mappings are already in the database.\n"
                        f"No new mappings were saved."
                    )
                    self._update_status("All mappings already present in database.")
                else:
                    messagebox.showinfo(
                        "No Required Mappings Found", 
                        "No required type mappings found to save.\n"
                        "Make sure columns are mapped to required types first."
                    )
                    self._update_status("No required mappings found.")
            if failed_count > 0:
                messagebox.showwarning(
                    "Some Mappings Failed", 
                    f"{failed_count} mappings failed to save. Check the logs for details."
                )
        return saved_count, failed_count, already_present_count

    def _on_row_review_click(self, event, sheet_name, tree, required_cols):
        region = tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        # Toggle validity
        is_valid = self.row_validity[sheet_name].get(idx, True)
        new_valid = not is_valid
        self.row_validity[sheet_name][idx] = new_valid
        # Update tag and status column
        tag = 'validrow' if new_valid else 'invalidrow'
        tree.item(row_id, tags=(tag,))
        vals = list(tree.item(row_id, 'values'))
        vals[-1] = "Valid" if new_valid else "Invalid"
        tree.item(row_id, values=vals)

    def _show_row_review(self, file_mapping, original_sheet_data=None):
        # Remove existing Row Review frame if present
        if self.row_review_frame:
            self.row_review_frame.destroy()
        
        # Hide all widgets/frames related to column mapping in the current tab
        tab = self.notebook.nametowidget(self.notebook.select())
        self._hidden_column_mapping_widgets = []
        for child in tab.winfo_children():
            # Hide everything except the Row Review frame (if present)
            if child != getattr(self, 'row_review_frame', None):
                child.grid_remove()
                self._hidden_column_mapping_widgets.append(child)
        
        # Add Row Review container
        self.row_review_frame = ttk.LabelFrame(tab, text="Row Review")
        self.row_review_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        self.row_review_frame.grid_rowconfigure(0, weight=1)
        self.row_review_frame.grid_columnconfigure(0, weight=1)
        
        # Add a notebook for row review per sheet
        self.row_review_notebook = ttk.Notebook(self.row_review_frame)
        self.row_review_notebook.grid(row=0, column=0, sticky=tk.NSEW)
        self.row_review_treeviews = {}
        self.row_validity = {}
        
        # Add a "Back to Column Mapping" button
        back_frame = ttk.Frame(self.row_review_frame)
        back_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
        back_frame.grid_columnconfigure(0, weight=1)
        back_btn = ttk.Button(back_frame, text="← Back to Column Mapping", 
                             command=lambda: self._show_column_mapping())
        back_btn.pack(side=tk.LEFT, padx=5)

        # Add a "Confirm Row Review" button below the Row Review window
        confirm_row_frame = ttk.Frame(tab)
        confirm_row_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
        confirm_row_frame.grid_columnconfigure(0, weight=1)
        confirm_row_btn = ttk.Button(confirm_row_frame, text="Confirm Row Review", command=self._on_confirm_row_review)
        confirm_row_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.confirm_row_frame = confirm_row_frame

        for sheet_idx, sheet in enumerate(file_mapping.sheets):
            frame = ttk.Frame(self.row_review_notebook)
            self.row_review_notebook.add(frame, text=sheet.sheet_name)
            # Use standard mapped column names for display and value extraction
            if self.column_mapper and hasattr(self.column_mapper, 'config'):
                required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
            else:
                required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
            
            # Define the correct column order for display
            display_column_order = ["code", "description", "unit", "quantity", "unit_price", "total_price"]
            
            # Use the display order, but only include columns that are actually mapped
            available_columns = []
            for col in display_column_order:
                if col in required_types:
                    available_columns.append(col)
            
            # Add any remaining required columns that weren't in the display order
            for col in required_types:
                if col not in available_columns:
                    available_columns.append(col)
            
            mapped_type_to_index = {}
            if hasattr(sheet, 'column_mappings'):
                for cm in sheet.column_mappings:
                    mapped_type_to_index[getattr(cm, 'mapped_type', None)] = cm.column_index
            
            # Use mapped type keys as columns (not uppercased)
            columns = ["#"] + available_columns + ["status"]
            tree = ttk.Treeview(frame, columns=columns, show="headings", height=12, selectmode="none")
            for col in columns:
                tree.heading(col, text=col.capitalize() if col != '#' else '#')
                if col == "status":
                    tree.column(col, width=80, anchor=tk.CENTER)
                else:
                    tree.column(col, width=120 if col != "#" else 40, anchor=tk.W)
            # Add scrollbars
            v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
            h_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
            tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            self.row_review_treeviews[sheet.sheet_name] = tree
            # Remove blue selection highlight
            style = ttk.Style(tree)
            style.map('Treeview', background=[('selected', '#FFEBEE')])  # Always red on select
            style.layout('Treeview.Item', [('Treeitem.padding', {'sticky': 'nswe', 'children': [('Treeitem.indicator', {'side': 'left', 'sticky': ''}), ('Treeitem.image', {'side': 'left', 'sticky': ''}), ('Treeitem.text', {'side': 'left', 'sticky': ''})]})])
            # Populate rows
            self.row_validity[sheet.sheet_name] = {}
            if hasattr(sheet, 'row_classifications'):
                for rc in sheet.row_classifications:
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                        except Exception:
                            row_data = None
                    if row_data is None:
                        row_data = []
                    row_values = [rc.row_index + 1]
                    for col in available_columns:
                        idx = mapped_type_to_index.get(col)
                        val = row_data[idx] if idx is not None and idx < len(row_data) else ""
                        
                        # Apply number formatting for specific columns
                        if col in ['unit_price', 'total_price']:
                            val = format_number(val, is_currency=True)
                        elif col == 'quantity':
                            val = format_number(val, is_currency=False)
                        
                        row_values.append(val)
                    is_valid = getattr(rc, 'row_type', None) == RowType.PRIMARY_LINE_ITEM or getattr(rc, 'row_type', None) == 'primary_line_item'
                    self.row_validity[sheet.sheet_name][rc.row_index] = is_valid
                    status = "Valid" if is_valid else "Invalid"
                    tag = 'validrow' if is_valid else 'invalidrow'
                    tree.insert('', 'end', iid=str(rc.row_index), values=row_values + [status], tags=(tag,))
            tree.tag_configure('validrow', background='#E8F5E9')  # light green
            tree.tag_configure('invalidrow', background='#FFEBEE')  # light red
            tree.bind('<Button-1>', lambda e, s=sheet.sheet_name, t=tree: self._on_row_review_click(e, s, t, required_types))
        # Optionally, select the first sheet by default
        if file_mapping.sheets:
            self.row_review_notebook.select(0)

    def _show_column_mapping(self):
        """Show the Column Mapping section and hide Row Review"""
        # Hide Row Review
        if self.row_review_frame:
            self.row_review_frame.destroy()
            self.row_review_frame = None
        # Restore all previously hidden column mapping widgets
        tab = self.notebook.nametowidget(self.notebook.select())
        if hasattr(self, '_hidden_column_mapping_widgets'):
            for widget in self._hidden_column_mapping_widgets:
                widget.grid()
            self._hidden_column_mapping_widgets = []

    def _on_confirm_row_review(self):
        """Handle row review confirmation and start categorization"""
        self._update_status("Row review confirmed. Starting categorization process...")
        # Get the current tab ID
        current_tab_id = self.notebook.select()
        file_mapping = self.tab_id_to_file_mapping.get(current_tab_id)
        if not file_mapping:
            messagebox.showerror("Error", "Could not find file mapping for categorization")
            return
        # --- NEW: Create normalized DataFrame for categorization with logging, only including valid rows ---
        try:
            import pandas as pd
            import logging
            logger = logging.getLogger(__name__)
            # Build a DataFrame from all sheets marked as BOQ, only including valid rows
            rows = []
            sheet_count = 0
            for sheet in getattr(file_mapping, 'sheets', []):
                if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                    continue
                sheet_count += 1
                # Use mapped_type as DataFrame columns
                col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]
                sheet_name = sheet.sheet_name
                validity_dict = self.row_validity.get(sheet_name, {})
                # For each row classification, get the row data
                for rc in getattr(sheet, 'row_classifications', []):
                    # Only include valid rows
                    if not validity_dict.get(rc.row_index, True):
                        continue
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                        except Exception:
                            row_data = None
                    if row_data is None:
                        row_data = []
                    # Build dict for DataFrame
                    row_dict = {col_headers[i]: row_data[i] if i < len(row_data) else '' for i in range(len(col_headers))}
                    row_dict['Source_Sheet'] = sheet.sheet_name
                    rows.append(row_dict)
            logger.info(f"Categorization: Processed {sheet_count} BOQ sheets, {len(rows)} valid rows for DataFrame.")
            if rows:
                df = pd.DataFrame(rows)
                # Ensure 'Description' column is present and capitalized
                if 'description' in df.columns and 'Description' not in df.columns:
                    df.rename(columns={'description': 'Description'}, inplace=True)
                logger.info(f"Categorization: DataFrame columns: {list(df.columns)}")
                logger.info(f"Categorization: First 3 rows: {df.head(3).to_dict(orient='records')}")
                logger.info(f"Categorization: First 10 Description values: {df['Description'].head(10).tolist() if 'Description' in df.columns else 'No Description column'}")
            else:
                df = None
                logger.warning("Categorization: No valid rows found for DataFrame.")
            file_mapping.dataframe = df
            if df is None or df.empty or 'Description' not in df.columns:
                messagebox.showerror("Categorization Error", "No valid data found for categorization. Please check that your sheets contain valid BOQ rows and that a 'Description' column is mapped.")
                return
        except Exception as e:
            import logging
            logging.getLogger(__name__).error(f"Failed to build DataFrame for categorization: {e}")
            file_mapping.dataframe = None
            messagebox.showerror("Categorization Error", f"Failed to build DataFrame for categorization: {e}")
            return
        # --- END NEW ---
        # Start categorization process
        self._start_categorization(file_mapping)
    
    def _start_categorization(self, file_mapping):
        """Start the categorization process"""
        if not CATEGORIZATION_AVAILABLE:
            messagebox.showerror("Error", "Categorization components not available")
            return
        
        try:
            # Show categorization dialog
            dialog = show_categorization_dialog(
                parent=self.root,
                controller=self.controller,
                file_mapping=file_mapping,
                on_complete=self._on_categorization_complete
            )
            
        except Exception as e:
            logger.error(f"Failed to start categorization: {e}")
            messagebox.showerror("Error", f"Failed to start categorization: {str(e)}")
    
    def _on_categorization_complete(self, final_dataframe, categorization_result):
        """Handle categorization completion"""
        try:
            current_tab_path = self.notebook.select()
            current_tab = self.notebook.select(current_tab_path)  # Get the actual tab widget
            print("[DEBUG] Current tab path:", current_tab_path)
            print("[DEBUG] Current tab widget:", current_tab)
            print("[DEBUG] File mapping tabs:", [str(file_data['file_mapping'].tab) for file_data in self.controller.current_files.values()])
            print("[DEBUG] Number of files:", len(self.controller.current_files))
            
            for file_key, file_data in self.controller.current_files.items():
                print("[DEBUG] Checking file_key:", file_key)
                print("[DEBUG] file_data['file_mapping'].tab:", file_data['file_mapping'].tab)
                print("[DEBUG] hasattr check:", hasattr(file_data['file_mapping'], 'tab'))
                print("[DEBUG] tab comparison:", str(file_data['file_mapping'].tab) == str(current_tab_path))
                
                if hasattr(file_data['file_mapping'], 'tab') and str(file_data['file_mapping'].tab) == str(current_tab_path):
                    print("[DEBUG] Found matching tab, storing data...")
                    # Store the categorized data
                    file_data['categorized_dataframe'] = final_dataframe
                    file_data['categorization_result'] = categorization_result
                    # Update the file mapping
                    file_mapping = file_data['file_mapping']
                    file_mapping.categorized_dataframe = final_dataframe
                    file_mapping.categorization_result = categorization_result
                    print("[DEBUG] About to call _show_final_categorized_data...")
                    # Show the final data grid in the main window - use the actual tab widget from file_mapping
                    self._show_final_categorized_data(file_mapping.tab, final_dataframe, categorization_result)
                    self._update_status("Categorization completed successfully - showing final data")
                    print("[DEBUG] _show_final_categorized_data call completed")
                    break
                else:
                    print("[DEBUG] Tab mismatch or no tab attribute")
        except Exception as e:
            print("[DEBUG] Exception in _on_categorization_complete:", e)
            import traceback
            traceback.print_exc()
            logger.error(f"Error handling categorization completion: {e}")
            messagebox.showerror("Error", f"Error handling categorization completion: {str(e)}")
    
    def _get_final_display_columns(self, file_mapping):
        # Get the mapped types in the order the user mapped them
        if hasattr(file_mapping, 'column_mappings'):
            mapped_types = [getattr(cm, 'mapped_type', None) for cm in file_mapping.column_mappings]
            # Remove None and duplicates, preserve order
            seen = set()
            ordered_types = []
            for t in mapped_types:
                if t and t not in seen:
                    ordered_types.append(t)
                    seen.add(t)
            # Add standard columns not mapped by user at the end
            std_order = ['code', 'sheet', 'category', 'description', 'quantity', 'unit_price', 'total_price', 'unit']
            for t in std_order:
                if t not in ordered_types:
                    ordered_types.append(t)
            return ordered_types
        else:
            return ['code', 'sheet', 'category', 'description', 'quantity', 'unit_price', 'total_price', 'unit']

    def _build_final_grid_dataframe(self, file_mapping):
        import pandas as pd
        rows = []
        # Determine required types and display order as in row review
        if self.column_mapper and hasattr(self.column_mapper, 'config'):
            required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
        else:
            required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
        # Row review display order
        display_column_order = ["code", "sheet", "category", "description", "unit", "quantity", "unit_price", "total_price"]
        
        # Helper function to parse numbers
        def parse_number(val):
            if isinstance(val, (int, float)):
                return float(val)
            if pd.isna(val):
                return 0.0
            s = str(val).replace('\u202f', '').replace(' ', '').replace(',', '.')
            try:
                return float(s)
            except Exception:
                return 0.0
        
        # Build mapping from mapped_type to column index for each sheet
        for sheet in getattr(file_mapping, 'sheets', []):
            if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                continue
            mapped_type_to_index = {}
            if hasattr(sheet, 'column_mappings'):
                for cm in sheet.column_mappings:
                    mapped_type_to_index[getattr(cm, 'mapped_type', None)] = cm.column_index
            validity_dict = self.row_validity.get(sheet.sheet_name, {})
            if hasattr(sheet, 'row_classifications'):
                for rc in sheet.row_classifications:
                    # Only include valid rows
                    is_valid = getattr(rc, 'row_type', None) == RowType.PRIMARY_LINE_ITEM or getattr(rc, 'row_type', None) == 'primary_line_item'
                    if not is_valid:
                        continue
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                        except Exception:
                            row_data = None
                    if row_data is None:
                        row_data = []
                    row_dict = {}
                    # Add columns in display order, fill with values if mapped
                    for col in display_column_order:
                        if col == 'sheet':
                            row_dict['sheet'] = sheet.sheet_name
                        elif col == 'category':
                            # PATCH: Fill from final_dataframe if available
                            category_value = ''
                            if hasattr(file_mapping, 'categorized_dataframe') and file_mapping.categorized_dataframe is not None:
                                df = file_mapping.categorized_dataframe
                                # Try to match by Description and Sheet if possible
                                desc = ''
                                idx = mapped_type_to_index.get('description')
                                if idx is not None and idx < len(row_data):
                                    desc = row_data[idx]
                                sheet_col = sheet.sheet_name
                                # Try to find the row in the categorized_dataframe
                                if 'Description' in df.columns:
                                    match = df[(df['Description'] == desc) & (df.get('Source_Sheet', sheet_col) == sheet_col)]
                                    if not match.empty and 'Category' in match.columns:
                                        category_value = match.iloc[0]['Category']
                                    elif not match.empty and 'category' in match.columns:
                                        category_value = match.iloc[0]['category']
                                elif 'category' in df.columns:
                                    match = df[df['category'] == desc]
                                    if not match.empty:
                                        category_value = match.iloc[0]['category']
                            # Fallback to rc.category if not found
                            if not category_value:
                                category_value = getattr(rc, 'category', '')
                            row_dict['category'] = category_value
                        else:
                            idx = mapped_type_to_index.get(col)
                            val = row_data[idx] if idx is not None and idx < len(row_data) else ""
                            # Parse numeric values
                            if col in ['quantity', 'unit_price', 'total_price']:
                                val = parse_number(val)
                            row_dict[col] = val
                    rows.append(row_dict)
        # Build DataFrame
        if rows:
            df = pd.DataFrame(rows)
            # Ensure numeric columns are properly typed
            for col in ['quantity', 'unit_price', 'total_price']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            # Only keep columns in display order that are present
            columns = [col for col in display_column_order if col in df.columns]
            df = df[columns]
        else:
            df = pd.DataFrame(columns=display_column_order)
        return df

    def _show_final_categorized_data(self, tab, final_dataframe, categorization_result):
        print("[DEBUG] _show_final_categorized_data called for tab:", tab)
        try:
            # Get file_mapping for this tab (if it exists - won't exist for loaded analyses)
            file_mapping = None
            for file_data in self.controller.current_files.values():
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == tab:
                    file_mapping = file_data['file_mapping']
                    break
            
            # For loaded analyses, use the DataFrame directly; for new analyses, build from file mapping
            if file_mapping is not None:
                # Build the DataFrame for the final grid using row review logic
                display_df = self._build_final_grid_dataframe(file_mapping)
            else:
                # For loaded analyses, use the provided DataFrame directly
                display_df = final_dataframe.copy()
                # Remove the 'category_internal' column if it exists (it's not needed for display)
                if 'category_internal' in display_df.columns:
                    display_df = display_df.drop('category_internal', axis=1)
                print(f"[DEBUG] Using loaded DataFrame directly with shape: {display_df.shape}")
            
            # Clear the current tab content
            for widget in tab.winfo_children():
                widget.destroy()
            # Create main frame
            main_frame = ttk.Frame(tab)
            main_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=10, pady=10)
            # Configure grid weights
            tab.grid_rowconfigure(0, weight=0)
            tab.grid_rowconfigure(1, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            main_frame.grid_rowconfigure(1, weight=1)
            main_frame.grid_columnconfigure(0, weight=1)
            # Title and instructions
            title_label = ttk.Label(main_frame, text="Final Categorized Data", font=("TkDefaultFont", 14, "bold"))
            title_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
            instructions = """
            Review the categorized data below. You can make corrections by double-clicking the category cell and selecting a value from the dropdown.\nChanges will be saved when you click 'Apply Changes', 'Summarize', 'Save Analysis', or 'Export Data'.
            """
            instruction_label = ttk.Label(main_frame, text=instructions, wraplength=800, justify=tk.LEFT)
            instruction_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
            # Create frame for the treeview
            tree_frame = ttk.Frame(main_frame)
            tree_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)
            # Only show columns that exist in the DataFrame
            final_display_columns = list(display_df.columns)
            tree = ttk.Treeview(tree_frame, columns=final_display_columns, show='headings', height=20)
            for col in final_display_columns:
                tree.heading(col, text=col.capitalize() if col != '#' else '#')
                tree.column(col, width=150, minwidth=100)
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.grid(row=0, column=0, sticky=tk.NSEW)
            vsb.grid(row=0, column=1, sticky=tk.NS)
            hsb.grid(row=1, column=0, sticky=tk.EW)
            # Set selection color to light blue and text color to black for readability
            style = ttk.Style(tree)
            style.map('Treeview', background=[('selected', '#B3E5FC')], foreground=[('selected', 'black')])
            self._populate_final_data_treeview(tree, display_df, final_display_columns)
            self._enable_final_data_editing(tree, display_df)
            # --- SUMMARY GRID PLACEHOLDER ---
            summary_frame = ttk.Frame(main_frame)
            summary_frame.grid(row=3, column=0, sticky=tk.EW, pady=(0, 10))
            summary_frame.grid_remove()  # Hide by default
            tab.summary_frame = summary_frame
            tab.summary_tree = None
            def show_summary_grid():
                import pandas as pd
                from core.manual_categorizer import get_manual_categorization_categories
                
                # Get pretty categories - this is now our single source of truth
                categories_pretty = get_manual_categorization_categories()
                
                # Helper to robustly parse numbers
                def parse_number(val):
                    if isinstance(val, (int, float)):
                        return float(val)
                    if pd.isna(val):
                        return 0.0
                    s = str(val).replace('\u202f', '').replace(' ', '').replace(',', '.')
                    try:
                        return float(s)
                    except Exception as e:
                        print(f"[DEBUG] Failed to parse number '{val}': {e}")
                        return 0.0
                
                # Get the current DataFrame (now uses pretty categories directly)
                df = tab.final_dataframe if hasattr(tab, 'final_dataframe') else display_df
                print("[DEBUG] DataFrame columns:", df.columns.tolist())
                print("[DEBUG] First few rows of DataFrame:")
                print(df.head().to_string())
                
                # Ensure numeric columns are properly parsed
                price_col = 'Total_price' if 'Total_price' in df.columns else 'total_price'
                print(f"[DEBUG] Using price column: {price_col}")
                if price_col in df.columns:
                    df[price_col] = df[price_col].apply(parse_number)
                    print(f"[DEBUG] Price column after parsing:")
                    print(df[price_col].head().to_string())
                
                # Group by category (now using pretty categories directly)
                cat_col = 'category'
                if cat_col in df.columns:
                    # Categories are already in pretty format, just group directly
                    print(f"[DEBUG] Unique categories before grouping:", df[cat_col].unique())
                    
                    # Create the summary dictionary with pretty category names
                    summary_dict = df.groupby(cat_col)[price_col].sum().to_dict()
                    print("[DEBUG] Summary dict after grouping:", summary_dict)
                    
                    # No need for complex mapping - categories are already pretty
                    # Just ensure we have all predefined categories with zero values if not present
                    final_summary = {}
                    for cat_pretty in categories_pretty:
                        final_summary[cat_pretty] = summary_dict.get(cat_pretty, 0.0)
                    
                    print("[DEBUG] Final summary:", final_summary)
                else:
                    print("[DEBUG] No category column found!")
                    final_summary = {cat: 0.0 for cat in categories_pretty}
                
                offer_label = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else 'Offer'
                summary_columns = ['Offer'] + categories_pretty
                
                # Use the final summary to get values in the correct order
                summary_values = [offer_label] + [final_summary[cat] for cat in categories_pretty]
                
                print("[DEBUG] Final summary values:", summary_values)
                
                # Remove old summary tree if present
                for widget in summary_frame.winfo_children():
                    widget.destroy()
                
                if len(summary_columns) <= 1:
                    summary_frame.grid_remove()
                    return
                
                summary_tree = ttk.Treeview(summary_frame, columns=summary_columns, show='headings', height=2)
                for col in summary_columns:
                    summary_tree.heading(col, text=col)
                    summary_tree.column(col, width=150, minwidth=100)
                
                # Format values for display
                display_values = []
                for i, val in enumerate(summary_values):
                    if i == 0:  # Offer label
                        display_values.append(str(val))
                    else:  # Numeric values
                        try:
                            num_val = float(val)
                            display_values.append(f"{num_val:,.2f}".replace(',', ' ').replace('.', ','))
                        except (ValueError, TypeError):
                            display_values.append(str(val))
                
                summary_tree.insert('', 'end', values=display_values, tags=('offer',))
                summary_tree.grid(row=0, column=0, sticky=tk.EW)
                summary_frame.grid()
                tab.summary_tree = summary_tree
                
                # --- CATEGORY FILTERING FEATURE ---
                tab._active_category_filter = None  # Track the current filter
                def on_summary_double_click(event):
                    region = summary_tree.identify('region', event.x, event.y)
                    if region != 'cell':
                        return
                    col_id = summary_tree.identify_column(event.x)
                    col_num = int(col_id[1:]) - 1  # 0-based
                    if col_num == 0:
                        return  # Ignore 'Offer' column
                    category_pretty = summary_columns[col_num]
                    
                    # Toggle filter: if already filtered to this, remove; else filter
                    if getattr(tab, '_active_category_filter', None) == category_pretty:
                        tab._active_category_filter = None
                        filtered_df = tab.final_dataframe if hasattr(tab, 'final_dataframe') else display_df
                    else:
                        tab._active_category_filter = category_pretty
                        df_full = tab.final_dataframe if hasattr(tab, 'final_dataframe') else display_df
                        # Filter by pretty category directly - no conversion needed
                        filtered_df = df_full[df_full['category'] == category_pretty]
                    
                    # Repopulate the main grid with the filtered DataFrame
                    self._populate_final_data_treeview(tab.final_data_tree, filtered_df, final_display_columns)
                    # Update the reference so further edits work on the filtered view
                    tab._filtered_dataframe = filtered_df
                
                summary_tree.bind('<Double-1>', on_summary_double_click)
                # --- END CATEGORY FILTERING FEATURE ---
            # --- END SUMMARY GRID PLACEHOLDER ---
            # Button frame at the bottom
            button_frame = ttk.Frame(main_frame)
            button_frame.grid(row=4, column=0, pady=(10, 0))
            summarize_button = ttk.Button(button_frame, text="Summarize", command=show_summary_grid)
            summarize_button.pack(side=tk.LEFT, padx=(0, 5))
            save_dataset_button = ttk.Button(button_frame, text="Save Dataset", command=lambda: self._save_dataset(tab))
            save_dataset_button.pack(side=tk.LEFT, padx=(0, 5))
            save_mappings_button = ttk.Button(button_frame, text="Save Mappings", command=lambda: self._save_mappings(tab))
            save_mappings_button.pack(side=tk.LEFT, padx=(0, 5))
            export_button = ttk.Button(button_frame, text="Export Data", 
                                      command=lambda: self._export_final_data(tab.final_dataframe, tab))
            export_button.pack(side=tk.LEFT, padx=(0, 5))
            tab.final_data_tree = tree
            tab.final_dataframe = display_df
            tab.categorization_result = categorization_result
        except Exception as e:
            print("[DEBUG] Exception in _show_final_categorized_data:", e)
            import traceback
            traceback.print_exc()
    
    def _populate_final_data_treeview(self, tree, dataframe, columns):
        """Populate the treeview with data from the final DataFrame, using pretty categories directly."""
        print(f"[DEBUG] _populate_final_data_treeview called with DataFrame shape: {dataframe.shape}")
        print(f"[DEBUG] DataFrame columns: {dataframe.columns.tolist()}")
        print(f"[DEBUG] Requested columns: {columns}")
        
        # Helper to format numbers consistently
        def format_number(val, is_currency=False):
            try:
                if pd.isna(val):
                    return ""
                if isinstance(val, str):
                    # Remove any existing formatting
                    val = val.replace(' ', '').replace('\u202f', '').replace(',', '.')
                num = float(val)
                return f"{num:,.2f}".replace(',', ' ').replace('.', ',')
            except (ValueError, TypeError):
                return str(val)
        
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        print(f"[DEBUG] Cleared existing items, now adding {len(dataframe)} rows")
        
        # Add data rows
        for index, row in dataframe.iterrows():
            values = []
            for col in columns:
                value = row.get(col, '')
                if pd.isna(value):
                    value = ''
                # Format based on column type
                if col == 'category':
                    # Category is already in pretty format - use directly
                    values.append(str(value))
                elif col in ['quantity', 'unit_price', 'total_price']:
                    # Format numeric columns
                    values.append(format_number(value))
                else:
                    values.append(str(value))
            
            print(f"[DEBUG] Adding row {index}: {values[:3]}...")  # Show first 3 values
            tree.insert('', 'end', values=values, tags=(f'row_{index}',))
        
        print(f"[DEBUG] Finished populating treeview with {len(dataframe)} rows")
    
    def _enable_final_data_editing(self, tree, dataframe):
        """Enable editing capabilities for the final data treeview. Only allow editing of the 'category' column with a dropdown."""
        from core.manual_categorizer import get_manual_categorization_categories
        categories_pretty = get_manual_categorization_categories()
        
        def on_double_click(event):
            try:
                row_id = tree.identify_row(event.y)
                column = tree.identify_column(event.x)
                if not row_id or not column:
                    return
                col_num = int(column[1:]) - 1  # Convert column identifier to index
                col_name = tree['columns'][col_num]
                if col_name != 'category':
                    return  # Only allow editing of the 'category' column
                current_values = tree.item(row_id, 'values')
                current_value = current_values[col_num] if col_num < len(current_values) else ''
                
                # Create combobox for category selection
                combo = ttk.Combobox(tree, values=categories_pretty, state='readonly')
                combo.set(current_value)
                bbox = tree.bbox(row_id, column)
                if bbox:
                    combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
                    def save_combo(event=None):
                        try:
                            new_pretty = combo.get()
                            values = list(tree.item(row_id, 'values'))
                            if col_num < len(values):
                                # Update display to pretty category
                                values[col_num] = new_pretty
                                tree.item(row_id, values=values)
                                # Update the underlying DataFrame with the pretty category (no conversion needed)
                                if dataframe is not None and row_id.isdigit():
                                    idx = int(row_id)
                                    if idx < len(dataframe):
                                        dataframe.at[idx, 'category'] = new_pretty
                            combo.destroy()
                        except Exception as e:
                            print(f"[DEBUG] Error saving combo: {e}")
                            combo.destroy()
                    def cancel_combo(event=None):
                        combo.destroy()
                    combo.bind('<Return>', save_combo)
                    combo.bind('<FocusOut>', save_combo)
                    combo.bind('<Escape>', cancel_combo)
                    combo.focus()
            except Exception as e:
                print(f"[DEBUG] Error in double-click editing: {e}")
        tree.bind('<Double-1>', on_double_click)
        print("[DEBUG] Double-click binding added to treeview (category only)")
    
    def _apply_final_data_changes(self, tree, dataframe, tab):
        """Apply changes made in the grid to the final DataFrame"""
        try:
            # Get all items from treeview
            items = tree.get_children()
            
            # Create a new DataFrame with the updated values
            updated_data = []
            columns = [tree.heading(col)['text'] for col in tree['columns']]
            
            for item in items:
                values = tree.item(item, 'values')
                row_data = dict(zip(columns, values))
                updated_data.append(row_data)
            
            # Update the DataFrame
            updated_dataframe = pd.DataFrame(updated_data)
            
            # Update stored references
            tab.final_dataframe = updated_dataframe
            
            # Update the file mapping
            current_tab = self.notebook.select()
            for file_key, file_data in self.controller.current_files.items():
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == current_tab:
                    file_mapping = file_data['file_mapping']
                    file_mapping.categorized_dataframe = updated_dataframe
                    file_data['categorized_dataframe'] = updated_dataframe
                    break
                    
            messagebox.showinfo("Success", "Changes applied successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply changes: {str(e)}")
    
    def _export_final_data(self, dataframe, tab=None):
        """Export the final categorized data to a nicely formatted Excel file with summary and validation."""
        try:
            import pandas as pd
            from tkinter import filedialog, messagebox
            import xlsxwriter
            import numpy as np
            from core.manual_categorizer import get_manual_categorization_categories
            
            # Get pretty categories - our single source of truth
            categories_pretty = get_manual_categorization_categories()
            
            # Helper to robustly parse numbers
            def parse_number(val):
                if isinstance(val, (int, float)):
                    return float(val)
                if pd.isna(val):
                    return 0.0
                s = str(val).replace('\u202f', '').replace(' ', '').replace(',', '.')
                try:
                    return float(s)
                except Exception:
                    return 0.0
            
            # Prompt for file path
            file_path = filedialog.asksaveasfilename(
                title="Export Categorized Data",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not file_path:
                return
            
            # Prepare main data
            df = dataframe.copy()
            
            # Categories are already in pretty format - no conversion needed
            print(f"[DEBUG] Exporting with categories: {df['category'].unique() if 'category' in df.columns else 'No category column'}")
            
            # Ensure numeric columns are numbers
            for col in ['quantity', 'unit_price', 'total_price']:
                if col in df.columns:
                    df[col] = df[col].apply(parse_number)
            
            # Write to Excel with formatting
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Data', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Data']
                
                # Format headers
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)
                
                # Format numeric columns
                num_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
                for col in ['quantity', 'unit_price', 'total_price']:
                    if col in df.columns:
                        col_idx = df.columns.get_loc(col)
                        worksheet.set_column(col_idx, col_idx, 15, num_format)
                
                # Data validation for category column
                if 'category' in df.columns:
                    cat_col_idx = df.columns.get_loc('category')
                    worksheet.data_validation(1, cat_col_idx, len(df), cat_col_idx, {
                        'validate': 'list',
                        'source': categories_pretty,
                        'input_message': 'Select a category from the list.'
                    })
                
                # Autofit columns
                for i, col in enumerate(df.columns):
                    maxlen = max(
                        [len(str(x)) for x in df[col].astype(str).values] + [len(str(col))]
                    )
                    worksheet.set_column(i, i, min(maxlen + 2, 30))
                
                # Add summary sheet
                if tab and hasattr(tab, 'summary_frame') and hasattr(tab, 'final_dataframe'):
                    df_summary = tab.final_dataframe.copy()
                    
                    # Ensure price column is numeric
                    price_col = 'total_price' if 'total_price' in df_summary.columns else 'Total_price'
                    if price_col in df_summary.columns:
                        df_summary[price_col] = df_summary[price_col].apply(parse_number)
                    
                    # Create summary by grouping pretty categories directly
                    if 'category' in df_summary.columns:
                        summary_dict = df_summary.groupby('category')[price_col].sum().to_dict()
                        
                        # Ensure all predefined categories are present
                        final_summary = {}
                        for cat_pretty in categories_pretty:
                            final_summary[cat_pretty] = summary_dict.get(cat_pretty, 0.0)
                    else:
                        final_summary = {cat: 0.0 for cat in categories_pretty}
                    
                    # Create summary sheet
                    summary_ws = workbook.add_worksheet('Summary')
                    
                    # Write headers with formatting
                    summary_columns = ['Offer'] + categories_pretty
                    for col_idx, col in enumerate(summary_columns):
                        summary_ws.write(0, col_idx, col, header_format)
                    
                    # Write values
                    offer_label = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else 'Offer'
                    summary_values = [offer_label] + [final_summary[cat] for cat in categories_pretty]
                    
                    # Write summary row with formatting
                    for col_idx, val in enumerate(summary_values):
                        if col_idx == 0:
                            summary_ws.write(1, col_idx, val)
                        else:
                            summary_ws.write_number(1, col_idx, val, num_format)
                    
                    # Format columns
                    for i in range(len(summary_columns)):
                        if i == 0:
                            summary_ws.set_column(i, i, 25)  # Offer column: text
                        else:
                            summary_ws.set_column(i, i, 18, num_format)  # Category columns: number format
                
                messagebox.showinfo("Success", f"Data exported to: {file_path}")
                self._update_status(f"Exported data to: {file_path}")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def run(self):
        """Start the main application loop."""
        self.root.mainloop()

    def _prompt_offer_name(self):
        # Simple dialog to prompt for offer name/label
        import tkinter.simpledialog
        offer_name = tkinter.simpledialog.askstring("Offer Name", "Enter a name or label for this offer:", parent=self.root)
        if offer_name is not None and offer_name.strip() != "":
            return offer_name.strip()
        return None

    def _save_dataset(self, tab):
        """Save the current DataFrame as a pickle file."""
        df = getattr(tab, 'final_dataframe', None)
        if df is None:
            messagebox.showerror("Error", "No dataset to save.")
            return
        
        print(f"[DEBUG] Saving DataFrame with shape: {df.shape}")
        print(f"[DEBUG] DataFrame columns: {df.columns.tolist()}")
        print(f"[DEBUG] DataFrame first few rows:")
        print(df.head().to_string())
        
        if df.empty:
            messagebox.showerror("Error", "Dataset is empty - nothing to save.")
            return
            
        # Get the current offer name
        offer_name = self.current_offer_name if hasattr(self, 'current_offer_name') else None
        
        # Create a dictionary with all necessary data
        save_data = {
            'dataframe': df,
            'offer_name': offer_name,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        file_path = filedialog.asksaveasfilename(
            title="Save Dataset as Pickle",
            defaultextension=".pkl",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'wb') as f:
                    pickle.dump(save_data, f)
                print(f"[DEBUG] Successfully saved to {file_path}")
                messagebox.showinfo("Success", f"Dataset saved to: {file_path}")
                self._update_status(f"Dataset saved to: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save dataset: {str(e)}")
                print(f"[DEBUG] Failed to save: {e}")
                import traceback
                traceback.print_exc()

    def _save_mappings(self, tab):
        """Save the current mappings (sheet, column, row) as a pickle file."""
        # Find the file_mapping for this tab
        file_mapping = None
        for file_data in self.controller.current_files.values():
            if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == tab:
                file_mapping = file_data['file_mapping']
                break
        if file_mapping is None:
            messagebox.showerror("Error", "No mappings to save.")
            return
        
        # Prepare mappings data
        mappings_data = {
            'sheets': [],
            'row_validity': self.row_validity.copy(),
        }
        
        # Save relevant info from each sheet
        for sheet in getattr(file_mapping, 'sheets', []):
            sheet_info = {
                'sheet_name': getattr(sheet, 'sheet_name', None),
                'sheet_type': getattr(sheet, 'sheet_type', None),
                'column_mappings': [],
                'row_classifications': [],
            }
            # Column mappings
            for cm in getattr(sheet, 'column_mappings', []):
                cm_dict = cm.__dict__.copy() if hasattr(cm, '__dict__') else dict(cm)
                sheet_info['column_mappings'].append(cm_dict)
            # Row classifications
            for rc in getattr(sheet, 'row_classifications', []):
                rc_dict = rc.__dict__.copy() if hasattr(rc, '__dict__') else dict(rc)
                sheet_info['row_classifications'].append(rc_dict)
            mappings_data['sheets'].append(sheet_info)
        
        # IMPORTANT: Save the final categorized DataFrame if it exists
        if hasattr(tab, 'final_dataframe') and tab.final_dataframe is not None:
            mappings_data['final_dataframe'] = tab.final_dataframe.copy()
            print(f"[DEBUG] Saving final DataFrame with shape: {tab.final_dataframe.shape}")
        
        # Save offer name if available
        if hasattr(self, 'current_offer_name') and self.current_offer_name:
            mappings_data['offer_name'] = self.current_offer_name
        
        # Optionally save column_mapper config if needed for re-use
        if hasattr(file_mapping, 'column_mapper'):
            try:
                mappings_data['column_mapper'] = file_mapping.column_mapper
            except Exception:
                pass
        
        file_path = filedialog.asksaveasfilename(
            title="Save Mappings as Pickle",
            defaultextension=".pkl",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'wb') as f:
                    pickle.dump(mappings_data, f)
                messagebox.showinfo("Success", f"Mappings saved to: {file_path}")
                self._update_status(f"Mappings saved to: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save mappings: {str(e)}")

    def _load_analysis(self):
        """Handle loading a previously saved analysis"""
        file_path = filedialog.askopenfilename(
            title="Load Analysis",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'rb') as f:
                    loaded_data = pickle.load(f)
                
                print(f"[DEBUG] Loaded data type: {type(loaded_data)}")
                
                # Handle both old format (just DataFrame) and new format (dictionary)
                if isinstance(loaded_data, dict):
                    df = loaded_data.get('dataframe')
                    self.current_offer_name = loaded_data.get('offer_name')
                    print(f"[DEBUG] Dictionary format - DataFrame shape: {df.shape if df is not None else 'None'}")
                    print(f"[DEBUG] Offer name: {self.current_offer_name}")
                else:
                    df = loaded_data
                    self.current_offer_name = None
                    print(f"[DEBUG] Direct DataFrame format - shape: {df.shape if df is not None else 'None'}")
                
                if isinstance(df, pd.DataFrame) and not df.empty:
                    print(f"[DEBUG] DataFrame columns: {df.columns.tolist()}")
                    print(f"[DEBUG] DataFrame first few rows:")
                    print(df.head().to_string())
                    
                    # Create a new tab for the loaded analysis
                    tab = ttk.Frame(self.notebook)
                    self.notebook.add(tab, text=f"Loaded Analysis {os.path.basename(file_path)}")
                    self.notebook.select(tab)
                    
                    # Show the data in the grid
                    self._show_final_categorized_data(tab, df, None)
                    self._update_status(f"Analysis loaded from: {file_path}")
                elif isinstance(df, pd.DataFrame) and df.empty:
                    messagebox.showerror("Error", "The loaded DataFrame is empty")
                    print("[DEBUG] DataFrame is empty")
                else:
                    messagebox.showerror("Error", "Invalid analysis file format - no DataFrame found")
                    print(f"[DEBUG] Invalid format - df type: {type(df)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load analysis: {str(e)}")
                logger.error(f"Failed to load analysis from {file_path}: {e}", exc_info=True)
                print(f"[DEBUG] Exception loading analysis: {e}")
                import traceback
                traceback.print_exc()

    def _use_mapping(self):
        """Handle loading and applying a previously saved mapping"""
        # Step 1: Load the mapping file
        mapping_path = filedialog.askopenfilename(
            title="Select Mapping File",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if not mapping_path:
            return
            
        try:
            with open(mapping_path, 'rb') as f:
                mapping_data = pickle.load(f)
            
            # Validate mapping data structure
            if not isinstance(mapping_data, dict) or 'sheets' not in mapping_data:
                messagebox.showerror("Error", "Invalid mapping file format")
                return
                
            self.saved_mapping = mapping_data
            print(f"[DEBUG] Loaded mapping with {len(mapping_data['sheets'])} sheets")
            
            # Debug: Check if mapping contains final categorized data
            if 'final_dataframe' in mapping_data:
                df = mapping_data['final_dataframe']
                print(f"[DEBUG] Mapping contains final DataFrame with shape: {df.shape if df is not None else None}")
                if df is not None and 'category' in df.columns:
                    print(f"[DEBUG] Final DataFrame has category column with {len(df)} rows")
                else:
                    print("[DEBUG] Final DataFrame missing category column")
            else:
                print("[DEBUG] Mapping does NOT contain final DataFrame - will require categorization")
            
            # Step 2: Prompt for BOQ file to analyze
            self._update_status("Mapping loaded. Please select a BOQ file to analyze.")
            
            # Prompt for offer name
            offer_name = self._prompt_offer_name()
            if offer_name is None:
                self._update_status("Analysis cancelled (no offer name provided).")
                return
            self.current_offer_name = offer_name
            
            # Select BOQ file
            filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            filename = filedialog.askopenfilename(title="Select BOQ File to Analyze", filetypes=filetypes)
            if not filename:
                self._update_status("Analysis cancelled (no file selected).")
                return
                
            # Step 3: Process the file with mapping validation
            self._process_file_with_mapping(filename, mapping_data)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load mapping: {str(e)}")
            print(f"[DEBUG] Error loading mapping: {e}")
            import traceback
            traceback.print_exc()

    def _process_file_with_mapping(self, filepath, mapping_data):
        """Process a BOQ file using a saved mapping"""
        self._update_status(f"Processing {os.path.basename(filepath)} with saved mapping...")
        
        # Create a new tab for the file
        filename = os.path.basename(filepath)
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=f"Mapped: {filename}")
        self.notebook.select(tab)
        
        # Configure grid for the tab frame
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        
        loading_label = ttk.Label(tab, text="Applying saved mapping...")
        loading_label.grid(row=0, column=0, pady=40, padx=100)
        self.root.update_idletasks()
        
        def process_with_mapping():
            try:
                # Step 1: Load the file and get sheets
                from core.file_processor import ExcelProcessor
                processor = ExcelProcessor()
                processor.load_file(Path(filepath))
                visible_sheets = processor.get_visible_sheets()
                
                print(f"[DEBUG] File has sheets: {visible_sheets}")
                print(f"[DEBUG] Mapping expects sheets: {[s['sheet_name'] for s in mapping_data['sheets']]}")
                
                # Step 2: Validate sheets
                mapping_sheets = {s['sheet_name'] for s in mapping_data['sheets']}
                file_sheets = set(visible_sheets)
                
                missing_sheets = mapping_sheets - file_sheets
                if missing_sheets:
                    error_msg = f"Missing sheets in file: {', '.join(missing_sheets)}"
                    self.root.after(0, lambda: messagebox.showerror("Sheet Validation Error", error_msg))
                    self.root.after(0, lambda: loading_label.destroy())
                    return
                
                # Step 3: Process the file (basic processing first)
                file_mapping = self.controller.process_file(
                    Path(filepath),
                    progress_callback=lambda p, m: self.root.after(0, self.update_progress, p, m),
                    sheet_filter=list(mapping_sheets),
                    sheet_types={sheet: "BOQ" for sheet in mapping_sheets}
                )
                
                # Step 4: Apply the saved mappings
                self.root.after(0, lambda: self._apply_saved_mappings(tab, file_mapping, mapping_data, loading_label))
                
            except Exception as e:
                print(f"[DEBUG] Error in process_with_mapping: {e}")
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: self._on_processing_error(tab, filename, loading_label))
        
        threading.Thread(target=process_with_mapping, daemon=True).start()

    def _apply_saved_mappings(self, tab, file_mapping, mapping_data, loading_widget):
        """Apply saved column and row mappings to the processed file with strict validation"""
        try:
            print("[DEBUG] Applying saved mappings with strict validation...")
            
            # Apply column mappings and validate structure
            for sheet in file_mapping.sheets:
                sheet_name = sheet.sheet_name
                
                # Find corresponding mapping
                saved_sheet = next((s for s in mapping_data['sheets'] if s['sheet_name'] == sheet_name), None)
                if not saved_sheet:
                    continue
                
                print(f"[DEBUG] Validating structure for sheet: {sheet_name}")
                
                # Strict validation: columns must match exactly
                if not self._validate_exact_column_structure(sheet, saved_sheet):
                    error_msg = (
                        f"Column structure mismatch in sheet '{sheet_name}'.\n\n"
                        "The new file does not have the exact same column structure as the saved mapping.\n"
                        "Mapping can only be applied to files with identical structure."
                    )
                    loading_widget.destroy()
                    messagebox.showerror("Structure Validation Failed", error_msg)
                    return
                
                # Strict validation: rows must match exactly
                if not self._validate_exact_row_structure(sheet, saved_sheet):
                    error_msg = (
                        f"Row structure mismatch in sheet '{sheet_name}'.\n\n"
                        "The new file does not have the exact same row content as the saved mapping.\n"
                        "Mapping can only be applied to files with identical structure and content."
                    )
                    loading_widget.destroy()
                    messagebox.showerror("Structure Validation Failed", error_msg)
                    return
                
                # Apply the mappings (since validation passed)
                self._apply_exact_column_mappings(sheet, saved_sheet)
                self._apply_exact_row_classifications(sheet, saved_sheet)
                
                print(f"[DEBUG] Sheet '{sheet_name}': Structure validated and mappings applied")
            
            # Store the file mapping and remove loading widget
            self.file_mapping = file_mapping
            self.column_mapper = file_mapping.column_mapper if hasattr(file_mapping, 'column_mapper') else None
            loading_widget.destroy()
            
            # Show Row Review directly (skip column mapping step)
            self._show_mapped_file_review(tab, file_mapping, mapping_data)
            
            success_msg = (
                f"Mapping applied successfully!\n\n"
                f"File structure matches perfectly with saved mapping.\n"
                f"All column mappings and row classifications have been applied."
            )
            self._update_status("Mapping applied successfully - identical structure confirmed")
            messagebox.showinfo("Mapping Applied", success_msg)
            
        except Exception as e:
            print(f"[DEBUG] Error applying mappings: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to apply mappings: {str(e)}")
            loading_widget.destroy()

    def _validate_exact_column_structure(self, sheet, saved_sheet):
        """Validate that the column structure matches exactly"""
        try:
            # Get current column headers
            current_headers = []
            if hasattr(sheet, 'column_mappings'):
                current_headers = [getattr(cm, 'original_header', '') for cm in sheet.column_mappings]
            
            # Get saved column headers
            saved_mappings = saved_sheet.get('column_mappings', [])
            saved_headers = []
            for saved_mapping in saved_mappings:
                if isinstance(saved_mapping, dict):
                    header = saved_mapping.get('original_header', '')
                    saved_headers.append(header)
            
            print(f"[DEBUG] Current headers ({len(current_headers)}): {current_headers}")
            print(f"[DEBUG] Saved headers ({len(saved_headers)}): {saved_headers}")
            
            # Must have same number of columns
            if len(current_headers) != len(saved_headers):
                print(f"[DEBUG] Column count mismatch: {len(current_headers)} vs {len(saved_headers)}")
                return False
            
            # Headers must match exactly (case-sensitive)
            for i, (current, saved) in enumerate(zip(current_headers, saved_headers)):
                if current.strip() != saved.strip():
                    print(f"[DEBUG] Header mismatch at position {i}: '{current}' vs '{saved}'")
                    return False
            
            print("[DEBUG] Column structure validation passed")
            return True
            
        except Exception as e:
            print(f"[DEBUG] Error validating column structure: {e}")
            return False

    def _validate_exact_row_structure(self, sheet, saved_sheet):
        """Validate that the row content matches exactly for valid rows only"""
        try:
            # Get saved row classifications - only the valid ones
            saved_classifications = saved_sheet.get('row_classifications', [])
            saved_valid_rows = []
            saved_valid_indices = []
            
            for saved_rc in saved_classifications:
                if isinstance(saved_rc, dict):
                    row_type = saved_rc.get('row_type', '')
                    # Only include rows that were marked as valid (primary line items)
                    if row_type in ['primary_line_item', 'PRIMARY_LINE_ITEM']:
                        row_data = saved_rc.get('row_data', [])
                        row_index = saved_rc.get('row_index', -1)
                        if row_data and row_index >= 0:
                            saved_valid_rows.append(row_data)
                            saved_valid_indices.append(row_index)
            
            print(f"[DEBUG] Saved mapping has {len(saved_valid_rows)} valid rows at indices: {saved_valid_indices[:10]}...")
            
            # Get current row data for the same indices
            current_valid_rows = []
            if hasattr(sheet, 'row_classifications'):
                for rc in sheet.row_classifications:
                    row_index = getattr(rc, 'row_index', -1)
                    
                    # Only check rows that were valid in the saved mapping
                    if row_index in saved_valid_indices:
                        row_data = getattr(rc, 'row_data', None)
                        if row_data is None and hasattr(sheet, 'sheet_data'):
                            try:
                                row_data = sheet.sheet_data[row_index]
                            except Exception:
                                row_data = []
                        if row_data is None:
                            row_data = []
                        current_valid_rows.append((row_index, row_data))
            
            # Sort current rows by index to match saved order
            current_valid_rows.sort(key=lambda x: x[0])
            current_row_data = [row_data for _, row_data in current_valid_rows]
            
            print(f"[DEBUG] Current file has {len(current_row_data)} rows at the expected valid indices")
            
            # Must have same number of valid rows
            if len(current_row_data) != len(saved_valid_rows):
                print(f"[DEBUG] Valid row count mismatch: {len(current_row_data)} vs {len(saved_valid_rows)}")
                print(f"[DEBUG] Missing indices: {set(saved_valid_indices) - {idx for idx, _ in current_valid_rows}}")
                return False
            
            # Each valid row must match exactly
            mismatched_rows = 0
            for i, (current_row, saved_row) in enumerate(zip(current_row_data, saved_valid_rows)):
                # Normalize for comparison (strip whitespace, handle empty cells)
                current_normalized = [str(cell).strip() if cell is not None else '' for cell in current_row]
                saved_normalized = [str(cell).strip() if cell is not None else '' for cell in saved_row]
                
                # Pad shorter row with empty strings
                max_len = max(len(current_normalized), len(saved_normalized))
                current_normalized.extend([''] * (max_len - len(current_normalized)))
                saved_normalized.extend([''] * (max_len - len(saved_normalized)))
                
                if current_normalized != saved_normalized:
                    mismatched_rows += 1
                    if mismatched_rows <= 3:  # Show first 3 mismatches for debugging
                        row_idx = saved_valid_indices[i] if i < len(saved_valid_indices) else i
                        print(f"[DEBUG] Valid row {row_idx} mismatch:")
                        print(f"  Current: {current_normalized[:3]}...")
                        print(f"  Saved:   {saved_normalized[:3]}...")
            
            if mismatched_rows > 0:
                print(f"[DEBUG] {mismatched_rows} valid rows don't match exactly")
                return False
            
            print("[DEBUG] Row structure validation passed - all valid rows match")
            return True
            
        except Exception as e:
            print(f"[DEBUG] Error validating row structure: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _apply_exact_column_mappings(self, sheet, saved_sheet):
        """Apply the saved column mappings exactly"""
        try:
            saved_mappings = saved_sheet.get('column_mappings', [])
            
            # Create mapping from header to saved mapping info
            header_to_mapping = {}
            for saved_mapping in saved_mappings:
                if isinstance(saved_mapping, dict):
                    header = saved_mapping.get('original_header', '')
                    header_to_mapping[header] = saved_mapping
            
            # Apply to current sheet
            for cm in sheet.column_mappings:
                original_header = getattr(cm, 'original_header', '')
                if original_header in header_to_mapping:
                    saved_mapping = header_to_mapping[original_header]
                    cm.mapped_type = saved_mapping.get('mapped_type', 'unknown')
                    cm.confidence = saved_mapping.get('confidence', 1.0)
                    cm.user_edited = saved_mapping.get('user_edited', True)
                    print(f"[DEBUG] Applied exact mapping: {original_header} -> {cm.mapped_type}")
            
        except Exception as e:
            print(f"[DEBUG] Error applying column mappings: {e}")

    def _apply_exact_row_classifications(self, sheet, saved_sheet):
        """Apply the saved row classifications exactly"""
        try:
            saved_classifications = saved_sheet.get('row_classifications', [])
            
            # Initialize row validity for this sheet
            sheet_name = sheet.sheet_name
            if not hasattr(self, 'row_validity'):
                self.row_validity = {}
            self.row_validity[sheet_name] = {}
            
            # Apply saved validity to corresponding rows
            if hasattr(sheet, 'row_classifications') and len(sheet.row_classifications) == len(saved_classifications):
                for i, (current_rc, saved_rc) in enumerate(zip(sheet.row_classifications, saved_classifications)):
                    if isinstance(saved_rc, dict):
                        saved_row_type = saved_rc.get('row_type', '')
                        is_valid = saved_row_type in ['primary_line_item', 'PRIMARY_LINE_ITEM']
                        
                        row_index = getattr(current_rc, 'row_index', i)
                        self.row_validity[sheet_name][row_index] = is_valid
                        
                        print(f"[DEBUG] Row {row_index}: {'Valid' if is_valid else 'Invalid'} (from saved mapping)")
            
            print(f"[DEBUG] Applied exact row classifications for {len(self.row_validity[sheet_name])} rows")
            
        except Exception as e:
            print(f"[DEBUG] Error applying row classifications: {e}")

    def _show_mapped_file_review(self, tab, file_mapping, mapping_data):
        """Show the file with applied mappings in Row Review mode"""
        try:
            # Clear tab content
            for widget in tab.winfo_children():
                widget.destroy()
            
            # Create main container frame
            tab_frame = ttk.Frame(tab)
            tab_frame.grid(row=0, column=0, sticky=tk.NSEW)
            
            # Configure grid layout
            tab_frame.grid_rowconfigure(0, weight=0)  # Info frame
            tab_frame.grid_rowconfigure(1, weight=1)  # Row review
            tab_frame.grid_rowconfigure(2, weight=0)  # Confirm button
            tab_frame.grid_columnconfigure(0, weight=1)
            
            # Add info frame
            info_frame = ttk.LabelFrame(tab_frame, text="Mapping Applied")
            info_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
            
            info_text = f"""Mapping successfully applied to {len(file_mapping.sheets)} sheets.
Column mappings and row validity have been pre-populated from the saved mapping.
Review the rows below and adjust validity as needed."""
            
            info_label = ttk.Label(info_frame, text=info_text, wraplength=800, justify=tk.LEFT)
            info_label.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
            
            # Show Row Review directly
            self._show_row_review_for_mapped_file(tab_frame, file_mapping)
            
            # Add confirm button
            confirm_frame = ttk.Frame(tab_frame)
            confirm_frame.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)
            confirm_frame.grid_columnconfigure(0, weight=1)
            
            confirm_btn = ttk.Button(confirm_frame, text="Confirm Row Review & Continue", 
                                   command=lambda: self._on_confirm_mapped_file_review(file_mapping))
            confirm_btn.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
            
            # Store file mapping reference
            file_mapping.tab = tab
            tab_id = str(tab)
            self.tab_id_to_file_mapping[tab_id] = file_mapping
            
        except Exception as e:
            print(f"[DEBUG] Error showing mapped file review: {e}")
            import traceback
            traceback.print_exc()

    def _show_row_review_for_mapped_file(self, parent_frame, file_mapping):
        """Show row review for a file with applied mappings"""
        # Create Row Review container
        row_review_frame = ttk.LabelFrame(parent_frame, text="Row Review (Pre-populated from Mapping)")
        row_review_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=5, pady=5)
        row_review_frame.grid_rowconfigure(0, weight=1)
        row_review_frame.grid_columnconfigure(0, weight=1)
        
        # Add notebook for sheets
        row_review_notebook = ttk.Notebook(row_review_frame)
        row_review_notebook.grid(row=0, column=0, sticky=tk.NSEW)
        
        # Initialize treeview storage
        if not hasattr(self, 'row_review_treeviews'):
            self.row_review_treeviews = {}
        
        # Create tabs for each sheet
        for sheet in file_mapping.sheets:
            self._create_row_review_tab_for_sheet(row_review_notebook, sheet)

    def _create_row_review_tab_for_sheet(self, notebook, sheet):
        """Create a row review tab for a specific sheet"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=sheet.sheet_name)
        
        # Determine required columns
        if self.column_mapper and hasattr(self.column_mapper, 'config'):
            required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
        else:
            required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
        
        # Display column order
        display_column_order = ["code", "description", "unit", "quantity", "unit_price", "total_price"]
        available_columns = [col for col in display_column_order if col in required_types]
        
        # Create column mapping
        mapped_type_to_index = {}
        if hasattr(sheet, 'column_mappings'):
            for cm in sheet.column_mappings:
                mapped_type_to_index[getattr(cm, 'mapped_type', None)] = cm.column_index
        
        # Create treeview
        columns = ["#"] + available_columns + ["status"]
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=12, selectmode="none")
        
        for col in columns:
            tree.heading(col, text=col.capitalize() if col != '#' else '#')
            if col == "status":
                tree.column(col, width=80, anchor=tk.CENTER)
            else:
                tree.column(col, width=120 if col != "#" else 40, anchor=tk.W)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Store treeview reference
        self.row_review_treeviews[sheet.sheet_name] = tree
        
        # Populate with data
        if hasattr(sheet, 'row_classifications'):
            for rc in sheet.row_classifications:
                row_data = getattr(rc, 'row_data', None)
                if row_data is None and hasattr(sheet, 'sheet_data'):
                    try:
                        row_data = sheet.sheet_data[rc.row_index]
                    except Exception:
                        row_data = []
                
                if row_data is None:
                    row_data = []
                
                # Build row values
                row_values = [rc.row_index + 1]
                for col in available_columns:
                    idx = mapped_type_to_index.get(col)
                    val = row_data[idx] if idx is not None and idx < len(row_data) else ""
                    
                    # Format numbers
                    if col in ['unit_price', 'total_price']:
                        val = format_number(val, is_currency=True)
                    elif col == 'quantity':
                        val = format_number(val, is_currency=False)
                    
                    row_values.append(val)
                
                # Get validity from pre-populated data
                is_valid = self.row_validity.get(sheet.sheet_name, {}).get(rc.row_index, False)
                status = "Valid" if is_valid else "Invalid"
                tag = 'validrow' if is_valid else 'invalidrow'
                
                tree.insert('', 'end', iid=str(rc.row_index), 
                          values=row_values + [status], tags=(tag,))
        
        # Configure tags
        tree.tag_configure('validrow', background='#E8F5E9')  # light green
        tree.tag_configure('invalidrow', background='#FFEBEE')  # light red
        
        # Bind click events
        tree.bind('<Button-1>', lambda e, s=sheet.sheet_name, t=tree: self._on_row_review_click(e, s, t, required_types))

    def _on_confirm_mapped_file_review(self, file_mapping):
        """Handle confirmation of row review for mapped file"""
        try:
            print(f"[DEBUG] _on_confirm_mapped_file_review called")
            print(f"[DEBUG] Has saved_mapping: {hasattr(self, 'saved_mapping')}")
            if hasattr(self, 'saved_mapping'):
                print(f"[DEBUG] Saved mapping keys: {list(self.saved_mapping.keys())}")
            
            # Check if we have saved categories from the mapping
            if hasattr(self, 'saved_mapping') and 'final_dataframe' in self.saved_mapping:
                print("[DEBUG] Found saved categories in mapping - applying directly without categorization")
                
                # Get the saved categorized DataFrame
                saved_df = self.saved_mapping['final_dataframe'].copy()
                
                # Build the current DataFrame structure (without categories)
                import pandas as pd
                rows = []
                sheet_count = 0
                
                for sheet in getattr(file_mapping, 'sheets', []):
                    if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                        continue
                        
                    sheet_count += 1
                    col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]
                    sheet_name = sheet.sheet_name
                    validity_dict = self.row_validity.get(sheet_name, {})
                    
                    for rc in getattr(sheet, 'row_classifications', []):
                        # Only include valid rows
                        if not validity_dict.get(rc.row_index, False):
                            continue
                            
                        row_data = getattr(rc, 'row_data', None)
                        if row_data is None and hasattr(sheet, 'sheet_data'):
                            try:
                                row_data = sheet.sheet_data[rc.row_index]
                            except Exception:
                                row_data = None
                        
                        if row_data is None:
                            row_data = []
                        
                        # Build dict for DataFrame
                        row_dict = {col_headers[i]: row_data[i] if i < len(row_data) else '' 
                                  for i in range(len(col_headers))}
                        row_dict['Source_Sheet'] = sheet.sheet_name
                        rows.append(row_dict)
                
                if rows:
                    current_df = pd.DataFrame(rows)
                    if 'description' in current_df.columns and 'Description' not in current_df.columns:
                        current_df.rename(columns={'description': 'Description'}, inplace=True)
                    
                    # Apply saved categories to current data by matching descriptions
                    print(f"[DEBUG] Applying saved categories to {len(current_df)} rows")
                    
                    # Create category mapping from saved DataFrame
                    if 'Description' in saved_df.columns and 'category' in saved_df.columns:
                        category_mapping = {}
                        for _, row in saved_df.iterrows():
                            desc = str(row['Description']).strip().lower()
                            category = str(row['category']).strip()
                            if desc and category:
                                category_mapping[desc] = category
                        
                        print(f"[DEBUG] Created category mapping with {len(category_mapping)} entries")
                        
                        # Apply categories to current DataFrame (categories are already in pretty format)
                        current_df['category'] = current_df['Description'].apply(
                            lambda x: category_mapping.get(str(x).strip().lower(), '')
                        )
                        
                        # Count successful matches
                        matched_count = (current_df['category'] != '').sum()
                        print(f"[DEBUG] Successfully matched {matched_count} out of {len(current_df)} rows")
                        
                        # Store the categorized DataFrame
                        file_mapping.dataframe = current_df
                        file_mapping.categorized_dataframe = current_df
                        
                        # Show final categorized data directly (skip categorization dialog)
                        current_tab = file_mapping.tab
                        self._show_final_categorized_data(current_tab, current_df, None)
                        
                        self._update_status(f"Applied saved categories to {matched_count} rows - categorization complete")
                        
                        success_msg = (
                            f"Categories applied successfully from saved mapping!\n\n"
                            f"Matched {matched_count} out of {len(current_df)} rows.\n"
                            f"Categorization completed without manual review."
                        )
                        messagebox.showinfo("Categories Applied", success_msg)
                        return
                    else:
                        print("[DEBUG] Saved DataFrame missing required columns for category mapping")
                else:
                    print("[DEBUG] No valid rows found for category application")
            
            # Fallback to normal categorization if no saved categories found
            print("[DEBUG] No saved categories found - proceeding with normal categorization")
            self._update_status("Row review confirmed. Starting categorization process...")
            
            # Build DataFrame for categorization (same as normal workflow)
            import pandas as pd
            rows = []
            sheet_count = 0
            
            for sheet in getattr(file_mapping, 'sheets', []):
                if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                    continue
                    
                sheet_count += 1
                col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]
                sheet_name = sheet.sheet_name
                validity_dict = self.row_validity.get(sheet_name, {})
                
                for rc in getattr(sheet, 'row_classifications', []):
                    # Only include valid rows
                    if not validity_dict.get(rc.row_index, False):
                        continue
                        
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                        except Exception:
                            row_data = None
                    
                    if row_data is None:
                        row_data = []
                    
                    # Build dict for DataFrame
                    row_dict = {col_headers[i]: row_data[i] if i < len(row_data) else '' 
                              for i in range(len(col_headers))}
                    row_dict['Source_Sheet'] = sheet.sheet_name
                    rows.append(row_dict)
            
            if rows:
                df = pd.DataFrame(rows)
                if 'description' in df.columns and 'Description' not in df.columns:
                    df.rename(columns={'description': 'Description'}, inplace=True)
                
                file_mapping.dataframe = df
                print(f"[DEBUG] Built DataFrame for categorization: {df.shape}")
                
                # Start categorization
                self._start_categorization(file_mapping)
            else:
                messagebox.showerror("Error", "No valid rows found for categorization")
                
        except Exception as e:
            print(f"[DEBUG] Error in mapped file review confirmation: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to continue with categorization: {str(e)}")
