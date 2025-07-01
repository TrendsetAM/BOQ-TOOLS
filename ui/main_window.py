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
            # Show confirmation message
            messagebox.showinfo(
                "Mapping Updated", 
                f"Column '{column_name}' has been mapped to '{new_type}' with 100% confidence.{learning_message}"
            )
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
        
        # Show column mapping summary
        messagebox.showinfo(
            "Column Mappings Saved",
            f"Saved {total_saved} new column mappings for all sheets.\n"
            f"Already present: {total_already}\n"
            f"Failed: {total_failed}\n\n"
            f"Now triggering row mapping with updated column mappings..."
        )
        self._update_status(f"Saved {total_saved} new column mappings. Triggering row mapping...")
        
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
                messagebox.showinfo(
                    "Row Mapping Complete",
                    f"Row mapping completed for {len(updated_sheets)} sheets:\n"
                    f"{', '.join(updated_sheets)}\n\n"
                    f"The sheet tabs have been updated with the new row classifications."
                )
                self._update_status(f"Row mapping completed for {len(updated_sheets)} sheets.")
            else:
                messagebox.showwarning(
                    "Row Mapping Skipped",
                    "Row mapping was skipped because original sheet data was not available.\n"
                    "Please reload the file to enable row mapping functionality."
                )
                self._update_status("Row mapping skipped - original data not available.")
                
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
            # Update the file mapping with categorized data
            current_tab = self.notebook.select()
            for file_key, file_data in self.controller.current_files.items():
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == current_tab:
                    # Store the categorized data
                    file_data['categorized_dataframe'] = final_dataframe
                    file_data['categorization_result'] = categorization_result
                    
                    # Update the file mapping
                    file_mapping = file_data['file_mapping']
                    file_mapping.categorized_dataframe = final_dataframe
                    file_mapping.categorization_result = categorization_result
                    
                    # Show success message
                    messagebox.showinfo("Success", 
                                      "Categorization completed successfully!\n"
                                      "You can now review categories or export the data.")
                    
                    # Add categorization buttons to the tab
                    self._add_categorization_buttons(current_tab, file_mapping)
                    
                    self._update_status("Categorization completed successfully")
                    break
                    
        except Exception as e:
            logger.error(f"Error handling categorization completion: {e}")
            messagebox.showerror("Error", f"Error handling categorization completion: {str(e)}")
    
    def _add_categorization_buttons(self, tab, file_mapping):
        """Add categorization action buttons to the tab"""
        # Find or create the categorization buttons frame
        categorization_frame = None
        
        # Look for existing categorization frame
        for widget in tab.winfo_children():
            if hasattr(widget, 'categorization_frame'):
                categorization_frame = widget
                break
        
        if not categorization_frame:
            # Create new categorization frame
            categorization_frame = ttk.LabelFrame(tab, text="Categorization Actions")
            categorization_frame.categorization_frame = True  # Mark it
            
            # Add it after the row review section
            categorization_frame.grid(row=5, column=0, sticky=tk.EW, padx=5, pady=5)
        
        # Clear existing buttons
        for widget in categorization_frame.winfo_children():
            widget.destroy()
        
        # Create buttons
        button_frame = ttk.Frame(categorization_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Review Categories button
        review_btn = ttk.Button(button_frame, text="Review Categories", 
                               command=lambda: self._review_categories(file_mapping))
        review_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Show Statistics button
        stats_btn = ttk.Button(button_frame, text="Show Statistics", 
                              command=lambda: self._show_categorization_stats(file_mapping))
        stats_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Export Categorized Data button
        export_btn = ttk.Button(button_frame, text="Export Categorized Data", 
                               command=lambda: self._export_categorized_data(file_mapping))
        export_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Re-run Categorization button
        rerun_btn = ttk.Button(button_frame, text="Re-run Categorization", 
                              command=lambda: self._start_categorization(file_mapping))
        rerun_btn.pack(side=tk.LEFT)
    
    def _review_categories(self, file_mapping):
        """Open category review dialog"""
        if not CATEGORIZATION_AVAILABLE:
            messagebox.showerror("Error", "Category review not available")
            return
        
        try:
            # Get the categorized dataframe
            dataframe = getattr(file_mapping, 'categorized_dataframe', None)
            if dataframe is None:
                messagebox.showwarning("Warning", "No categorized data available for review")
                return
            
            # Show category review dialog
            dialog = show_category_review_dialog(
                parent=self.root,
                dataframe=dataframe,
                on_save=self._on_category_review_save
            )
            
        except Exception as e:
            logger.error(f"Failed to open category review: {e}")
            messagebox.showerror("Error", f"Failed to open category review: {str(e)}")
    
    def _on_category_review_save(self, updated_dataframe):
        """Handle category review save"""
        try:
            # Update the file mapping with the modified dataframe
            current_tab = self.notebook.select()
            for file_key, file_data in self.controller.current_files.items():
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == current_tab:
                    file_mapping = file_data['file_mapping']
                    file_mapping.categorized_dataframe = updated_dataframe
                    file_data['categorized_dataframe'] = updated_dataframe
                    
                    messagebox.showinfo("Success", "Category changes saved successfully!")
                    self._update_status("Category changes saved")
                    break
                    
        except Exception as e:
            logger.error(f"Error saving category changes: {e}")
            messagebox.showerror("Error", f"Error saving category changes: {str(e)}")
    
    def _show_categorization_stats(self, file_mapping):
        """Show categorization statistics dialog"""
        if not CATEGORIZATION_AVAILABLE:
            messagebox.showerror("Error", "Statistics dialog not available")
            return
        
        try:
            # Get the categorized dataframe
            dataframe = getattr(file_mapping, 'categorized_dataframe', None)
            if dataframe is None:
                messagebox.showwarning("Warning", "No categorized data available for statistics")
                return
            
            # Get categorization result
            categorization_result = getattr(file_mapping, 'categorization_result', None)
            
            # Show statistics dialog
            dialog = show_categorization_stats_dialog(
                parent=self.root,
                dataframe=dataframe,
                categorization_result=categorization_result
            )
            
        except Exception as e:
            logger.error(f"Failed to show statistics: {e}")
            messagebox.showerror("Error", f"Failed to show statistics: {str(e)}")
    
    def _export_categorized_data(self, file_mapping):
        """Export categorized data"""
        try:
            # Get the categorized dataframe
            dataframe = getattr(file_mapping, 'categorized_dataframe', None)
            if dataframe is None:
                messagebox.showwarning("Warning", "No categorized data available for export")
                return
            
            # Use the existing export functionality
            from tkinter import filedialog
            
            file_path = filedialog.asksaveasfilename(
                title="Export Categorized Data",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if file_path:
                if file_path.endswith('.csv'):
                    dataframe.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    dataframe.to_excel(file_path, index=False)
                else:
                    dataframe.to_csv(file_path, index=False)
                
                messagebox.showinfo("Success", f"Categorized data exported to: {file_path}")
                self._update_status(f"Exported categorized data to: {file_path}")
                
        except Exception as e:
            logger.error(f"Failed to export categorized data: {e}")
            messagebox.showerror("Error", f"Failed to export categorized data: {str(e)}")

    def run(self):
        """Start the main application loop."""
        self.root.mainloop()
