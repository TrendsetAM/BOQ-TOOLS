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

# Custom exceptions for validation failures
class ValidationError(Exception):
    """Custom exception for validation failures"""
    def __init__(self, message, validation_result=None):
        super().__init__(message)
        self.validation_result = validation_result

class PositionValidationError(ValidationError):
    """Specific exception for position-description validation failures"""
    pass

def _format_validation_error_message(validation_result, context=""):
    """
    Standardized error message formatting for position-description validation failures
    
    Args:
        validation_result: Validation result dictionary with errors and summary
        context: Additional context string (e.g., "Use Mapping", "Compare Full")
        
    Returns:
        str: Formatted error message for user display
    """
    if not validation_result or validation_result.get('is_valid', True):
        return "No validation errors found."
    
    # Base error message
    title = f"{context} - Structure Validation Failed" if context else "Structure Validation Failed"
    
    # Summary
    summary = validation_result.get('summary', 'Validation failed with unknown errors.')
    
    # Error details (limit to first 5 errors for readability)
    errors = validation_result.get('errors', [])
    error_details = "\n".join(errors[:5])
    if len(errors) > 5:
        error_details += f"\n... and {len(errors) - 5} more errors"
    
    # Mismatched positions for additional context
    mismatched_positions = validation_result.get('mismatched_positions', [])
    position_count = len(mismatched_positions)
    
    # Build comprehensive error message
    if context == "Use Mapping":
        action_guidance = (
            "This means the current file has a different structure than the saved mapping.\n"
            "You cannot apply this mapping to the current file.\n\n"
            "Possible solutions:\n"
            "• Use a file with the same structure as the original\n"
            "• Create a new mapping for this file structure\n"
            "• Verify you selected the correct mapping file"
        )
    elif context == "Compare Full":
        action_guidance = (
            "This means the comparison file has a different structure than the master BOQ.\n"
            "You cannot compare BOQs with different structures.\n\n"
            "Possible solutions:\n"
            "• Use a comparison file with the same structure\n"
            "• Create separate analyses for files with different structures\n"
            "• Verify you selected the correct comparison file"
        )
    else:
        action_guidance = (
            "The files have different structures and cannot be processed together.\n\n"
            "Please verify that you are using compatible files."
        )
    
    formatted_message = (
        f"{summary}\n\n"
        f"Details ({position_count} mismatches found):\n{error_details}\n\n"
        f"{action_guidance}"
    )
    
    return formatted_message

def _log_validation_failure(validation_result, context="", operation="validation"):
    """
    Standardized logging for validation failures
    
    Args:
        validation_result: Validation result dictionary
        context: Operation context (e.g., "Use Mapping", "Compare Full")
        operation: Type of operation being performed
    """
    if not validation_result:
        logger.error(f"{context} {operation}: No validation result provided")
        return
    
    if validation_result.get('is_valid', True):
        logger.info(f"{context} {operation}: Validation passed successfully")
        return
    
    # Log summary
    summary = validation_result.get('summary', 'Unknown validation failure')
    logger.error(f"{context} {operation} failed: {summary}")
    
    # Log detailed errors
    errors = validation_result.get('errors', [])
    logger.error(f"{context} {operation}: Found {len(errors)} validation errors:")
    for i, error in enumerate(errors[:10], 1):  # Log first 10 errors
        logger.error(f"  {i}. {error}")
    
    if len(errors) > 10:
        logger.error(f"  ... and {len(errors) - 10} more errors")
    
    # Log mismatched positions summary
    mismatched_positions = validation_result.get('mismatched_positions', [])
    if mismatched_positions:
        logger.debug(f"{context} {operation}: Mismatched positions details:")
        for position_info in mismatched_positions[:5]:  # Log first 5 for debugging
            if 'expected_description' in position_info:
                logger.debug(f"  Position {position_info.get('position', 'unknown')}: "
                           f"expected '{position_info.get('expected_description', '')}', "
                           f"got '{position_info.get('actual_description', '')}'")
            else:
                logger.debug(f"  Missing position {position_info.get('position', 'unknown')}: "
                           f"expected '{position_info.get('expected_description', '')}'")

def _handle_validation_failure(validation_result, context="", operation="validation", show_dialog=True):
    """
    Standardized validation failure handling with logging and user notification
    
    Args:
        validation_result: Validation result dictionary
        context: Operation context for error messages
        operation: Type of operation being performed
        show_dialog: Whether to show error dialog to user
        
    Returns:
        bool: False (indicating failure)
        
    Raises:
        PositionValidationError: Always raises this exception to terminate processing
    """
    # Log the failure
    _log_validation_failure(validation_result, context, operation)
    
    # Format user message
    error_message = _format_validation_error_message(validation_result, context)
    
    # Show dialog if requested
    if show_dialog:
        dialog_title = f"{context} - Structure Validation Failed" if context else "Structure Validation Failed"
        messagebox.showerror(dialog_title, error_message)
    
    # Log termination
    logger.warning(f"{context} {operation}: Process terminated due to validation failure")
    
    # Raise exception to terminate processing
    raise PositionValidationError(error_message, validation_result)

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

def format_number_eu(val):
    """Format a number with point as thousands separator and comma as decimal separator (e.g., 1.234.567,89)"""
    try:
        if val is None or val == '' or (isinstance(val, float) and (val != val)):
            return ''
        num = float(str(val).replace(' ', '').replace('\u202f', '').replace(',', '.'))
        # Format with US locale, then replace separators
        s = f"{num:,.2f}"
        # s is like '1,234,567.89' -> want '1.234.567,89'
        s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
        return s
    except Exception:
        return str(val)

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
        # Store comprehensive offer information
        self.current_offer_info = None
        self.previous_offer_info = None  # For subsequent BOQs in comparison
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
        file_menu.add_command(label="Clear All Files", command=self._clear_all_files)
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
        # Prompt for comprehensive offer information before opening file dialog
        offer_info = self._prompt_offer_info(is_first_boq=True)
        if offer_info is None:
            self._update_status("File open cancelled (no offer information provided).")
            return
        
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filenames = filedialog.askopenfilenames(title="Open Excel File", filetypes=filetypes)
        for file in filenames:
            # Check if there is already a file loaded in the current tab
            current_tab_id = self.notebook.select()
            if current_tab_id:
                current_tab = self.notebook.nametowidget(current_tab_id)
                # If the tab has a final_dataframe, trigger comparison logic
                if hasattr(current_tab, 'final_dataframe') and getattr(current_tab, 'final_dataframe', None) is not None:
                    self._compare_full(current_tab)
                    return
            # Otherwise, just open as new analysis
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
        print(f"[DEBUG] _on_processing_complete called for file: {filepath}")
        
        # Store the file mapping and column mapper
        self.file_mapping = file_mapping
        self.column_mapper = file_mapping.column_mapper if hasattr(file_mapping, 'column_mapper') else None
        
        # Store offer info in the controller's current_files for summary data collection
        file_key = str(Path(filepath).resolve())
        if file_key in self.controller.current_files:
            current_offer_info = getattr(self, 'current_offer_info', {})
            current_offer_name = getattr(self, 'current_offer_name', 'Unknown')
            
            # Enhanced offer info creation with better fallbacks
            offer_info = {
                'supplier_name': current_offer_info.get('supplier_name', current_offer_name),
                'project_name': current_offer_info.get('project_name', 'Unknown'),
                'date': current_offer_info.get('date', 'Unknown')
            }
            
            # Store under dynamic offer name for comparison datasets
            offer_name = offer_info['supplier_name']
            if 'offers' not in self.controller.current_files[file_key]:
                self.controller.current_files[file_key]['offers'] = {}
            self.controller.current_files[file_key]['offers'][offer_name] = offer_info
            # For backward compatibility, also store the last offer as 'offer_info'
            self.controller.current_files[file_key]['offer_info'] = offer_info
            print(f"[DEBUG] Stored offer info for offer_name '{offer_name}': {offer_info}")
            print(f"[DEBUG] Current offer info state: current_offer_info={current_offer_info}, current_offer_name={current_offer_name}")
        
        # Remove loading widget and populate tab
        loading_widget.destroy()
        self._populate_file_tab(tab, file_mapping)
        
        # Use centralized refresh method
        print(f"[DEBUG] Calling centralized summary grid refresh")
        self._refresh_summary_grid_centralized()
        
        # Update status
        self._update_status(f"Processing complete: {os.path.basename(filepath)}")

    def _on_processing_error(self, tab, filename, loading_widget):
        """Callback for when file processing fails. Runs in the main UI thread."""
        print(f"[DEBUG] _on_processing_error called for file: {filename}")
        loading_widget.destroy()
        # Use grid for consistency
        error_label = ttk.Label(tab, text=f"Failed to process {filename}.\nSee logs for details.", foreground="red")
        error_label.grid(row=0, column=0, pady=40, padx=100)
        self._update_status(f"Error processing {filename}")
        self.progress_var.set(0)
        # Refresh summary grid even on error to show current state
        print(f"[DEBUG] Calling centralized summary grid refresh after error")
        self._refresh_summary_grid_centralized()

    def _populate_file_tab(self, tab, file_mapping):
        # print("[DEBUG] _populate_file_tab called for tab:", tab)
        """Populates a tab with the processed data from a file mapping."""
        # Debug: print all sheet names and their types
        # print('DEBUG: Sheets in file_mapping:')
        for s in file_mapping.sheets:
            print(f'  {s.sheet_name} (sheet_type={getattr(s, "sheet_type", None)})')
        
        # Clear any existing widgets (like loading/error labels)
        for widget in tab.winfo_children():
            widget.destroy()
        
        # Create main container frame
        tab_frame = ttk.Frame(tab)
        tab_frame.grid(row=0, column=0, sticky=tk.NSEW)
        
        # Configure tab_frame's grid layout
        tab_frame.grid_rowconfigure(0, weight=1)  # For sheet_notebook (main content, expands vertically)
        tab_frame.grid_rowconfigure(1, weight=0)  # For confirm_col_btn
        tab_frame.grid_rowconfigure(2, weight=1)  # For row_review_container (expands vertically)
        tab_frame.grid_columnconfigure(0, weight=1)  # Only one column, expands horizontally

        # Create sheet notebook for individual sheet tabs
        sheet_notebook = ttk.Notebook(tab_frame)
        sheet_notebook.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        
        # Populate each sheet as a tab in the sheet_notebook
        for sheet in file_mapping.sheets:
            sheet_frame = ttk.Frame(sheet_notebook)
            sheet_notebook.add(sheet_frame, text=sheet.sheet_name)
            self._populate_sheet_tab(sheet_frame, sheet)
        
        # Add confirmation button for column mappings
        confirm_frame = ttk.Frame(tab_frame)
        confirm_frame.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
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
        sheet_frame.grid_rowconfigure(0, weight=0)  # Header row control
        sheet_frame.grid_rowconfigure(1, weight=1)  # Column mappings table
        sheet_frame.grid_columnconfigure(0, weight=1)
        
        # Header Row Control
        header_control_frame = ttk.Frame(sheet_frame)
        header_control_frame.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        
        # Get current header row (convert from 0-based to 1-based for display)
        current_header_row = getattr(sheet, 'header_row_index', 0) + 1
        header_row_var = tk.IntVar(value=current_header_row)
        
        # Get sheet data length for validation (use a reasonable default)
        # We'll load the file only when the user actually clicks the +/- buttons
        sheet_data_length = 50  # Reasonable default for most Excel files
        
        # Header row control widgets
        ttk.Label(header_control_frame, text="Header Row:").pack(side=tk.LEFT, padx=(0, 5))
        
        decrease_btn = ttk.Button(header_control_frame, text="−", width=3,
                                 command=lambda: self._on_header_row_decrease(sheet, header_row_var, entry_widget, 
                                                                             decrease_btn, increase_btn))
        decrease_btn.pack(side=tk.LEFT, padx=(0, 2))
        
        entry_widget = ttk.Entry(header_control_frame, textvariable=header_row_var, width=4, 
                                justify=tk.CENTER, state='readonly')
        entry_widget.pack(side=tk.LEFT, padx=(0, 2))
        
        increase_btn = ttk.Button(header_control_frame, text="+", width=3,
                                 command=lambda: self._on_header_row_increase(sheet, header_row_var, entry_widget,
                                                                             decrease_btn, increase_btn))
        increase_btn.pack(side=tk.LEFT)
        
        # Update button states based on current row
        self._update_header_row_buttons_simple(current_header_row, decrease_btn, increase_btn)
        
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
        
        # Required types (base required + successfully mapped new columns)
        if self.column_mapper and hasattr(self.column_mapper, 'config'):
            base_required_types = {col_type.value for col_type in self.column_mapper.config.get_required_columns()}
        else:
            base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        
        # Add new columns as "required" if they are successfully mapped (confidence > 0)
        required_types = base_required_types.copy()
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                confidence = getattr(mapping, 'confidence', 0)
                # Treat scope, manhours, wage as required if successfully mapped
                if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                    required_types.add(mapped_type)
        
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
    
    def _update_header_row_buttons(self, current_row, decrease_btn, increase_btn, max_rows):
        """Update the state of header row buttons based on current row"""
        # Disable decrease button if at minimum (row 1)
        if current_row <= 1:
            decrease_btn.config(state='disabled')
        else:
            decrease_btn.config(state='normal')
        
        # Disable increase button if at maximum
        if current_row >= max_rows:
            increase_btn.config(state='disabled')
        else:
            increase_btn.config(state='normal')
    
    def _update_header_row_buttons_simple(self, current_row, decrease_btn, increase_btn):
        """Update the state of header row buttons based on current row (simple version)"""
        # Disable decrease button if at minimum (row 1)
        if current_row <= 1:
            decrease_btn.config(state='disabled')
        else:
            decrease_btn.config(state='normal')
        
        # Increase button is always enabled (validation happens in reprocess method)
        increase_btn.config(state='normal')
    
    def _on_header_row_decrease(self, sheet, header_row_var, entry_widget, decrease_btn, increase_btn):
        """Decrease header row number and immediately refresh"""
        current_row = header_row_var.get()
        if current_row > 1:  # Minimum row 1
            new_row = current_row - 1
            header_row_var.set(new_row)
            
            # Update entry widget
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, str(new_row))
            entry_widget.config(state='readonly')
            
            # Immediately reprocess with new header row (validation happens inside)
            success = self._reprocess_sheet_with_header_row(sheet, new_row - 1)  # Convert to 0-based
            
            # Update button states based on success and new row
            if success:
                self._update_header_row_buttons_simple(new_row, decrease_btn, increase_btn)
            else:
                # Revert on failure
                header_row_var.set(current_row)
                entry_widget.config(state='normal')
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, str(current_row))
                entry_widget.config(state='readonly')
    
    def _on_header_row_increase(self, sheet, header_row_var, entry_widget, decrease_btn, increase_btn):
        """Increase header row number and immediately refresh"""
        current_row = header_row_var.get()
        new_row = current_row + 1
        header_row_var.set(new_row)
        
        # Update entry widget
        entry_widget.config(state='normal')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, str(new_row))
        entry_widget.config(state='readonly')
        
        # Immediately reprocess with new header row (validation happens inside)
        success = self._reprocess_sheet_with_header_row(sheet, new_row - 1)  # Convert to 0-based
        
        # Update button states based on success and new row
        if success:
            self._update_header_row_buttons_simple(new_row, decrease_btn, increase_btn)
        else:
            # Revert on failure
            header_row_var.set(current_row)
            entry_widget.config(state='normal')
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, str(current_row))
            entry_widget.config(state='readonly')
    
    def _reprocess_sheet_with_header_row(self, sheet, new_header_row_index):
        """Reprocess sheet with new header row and update UI"""
        try:
            # Get the current file path from the file mapping
            if not hasattr(self, 'file_mapping') or not self.file_mapping:
                messagebox.showerror("Error", "No file mapping available")
                return False
            
            # Get the file path from the metadata
            file_path = Path(self.file_mapping.metadata.file_path)
            if not file_path.exists():
                messagebox.showerror("Error", f"Original file not found: {file_path}")
                return False
            
            # Reload the file temporarily to get sheet data
            try:
                from core.file_processor import ExcelProcessor
                temp_processor = ExcelProcessor()
                temp_processor.load_file(file_path)
                
                # Get sheet data for the specific sheet
                sheet_data = temp_processor.get_sheet_data(sheet.sheet_name)
                
                if not sheet_data:
                    messagebox.showerror("Error", "No sheet data available for reprocessing")
                    return False
                
                # Validate the new header row index
                if new_header_row_index < 0 or new_header_row_index >= len(sheet_data):
                    messagebox.showerror("Error", f"Invalid header row {new_header_row_index + 1}. Sheet has {len(sheet_data)} rows.")
                    return False
                
            except Exception as e:
                logger.error(f"Failed to reload file and get sheet data: {e}")
                messagebox.showerror("Error", f"Failed to reload file: {str(e)}")
                return False
            
            # Use the column mapper to reprocess with forced header row
            if not self.column_mapper:
                messagebox.showerror("Error", "Column mapper not available")
                return False
            
            # Process with forced header row
            mapping_result = self.column_mapper.process_sheet_mapping_with_forced_header(
                sheet_data, new_header_row_index
            )
            
            # Update sheet object with new mapping results
            sheet.header_row_index = new_header_row_index
            sheet.confidence = mapping_result.overall_confidence
            
            # Convert mappings to the format expected by the sheet
            new_column_mappings = []
            for mapping in mapping_result.mappings:
                # Create a simple object with the required attributes
                col_mapping = type('ColumnMapping', (), {})()
                col_mapping.column_index = mapping.column_index
                col_mapping.original_header = mapping.original_header
                col_mapping.mapped_type = mapping.mapped_type.value
                col_mapping.confidence = mapping.confidence
                col_mapping.user_edited = True  # Mark as user-edited since user manually changed header row
                col_mapping.reasoning = mapping.reasoning
                new_column_mappings.append(col_mapping)
            
            sheet.column_mappings = new_column_mappings
            
            # CRITICAL FIX: Update the cached sheet data in the controller to reflect the new header row
            # This prevents data corruption during row classification
            if hasattr(self, 'controller') and self.controller:
                # Find the file key by matching the file mapping
                file_key = None
                for key, file_data in self.controller.current_files.items():
                    if file_data.get('file_mapping') == self.file_mapping:
                        file_key = key
                        break
                
                if file_key and hasattr(self.controller, 'current_files') and file_key in self.controller.current_files:
                    processor_results = self.controller.current_files[file_key].get('processor_results', {})
                    original_sheet_data = processor_results.get('sheet_data', {})
                    
                    # Update the cached sheet data with the new header row information
                    if sheet.sheet_name in original_sheet_data:
                        # Store the new header row index as metadata for this sheet
                        # This will be used during row classification to skip the correct header row
                        if 'header_row_indices' not in processor_results:
                            processor_results['header_row_indices'] = {}
                        processor_results['header_row_indices'][sheet.sheet_name] = new_header_row_index
                        
                        print(f"[DEBUG] Updated cached header row index for sheet '{sheet.sheet_name}' to {new_header_row_index}")
            
            # CRITICAL FIX: Set the sheet.sheet_data to the FULL original data
            # The row classifications contain indices that reference the original data structure
            # We should NOT remove the header row from sheet.sheet_data because the row indices
            # in the classifications are adjusted to account for the header row position
            sheet.sheet_data = sheet_data
            print(f"[DEBUG] Set sheet.sheet_data for '{sheet.sheet_name}' with full original data")
            print(f"[DEBUG] Data length: {len(sheet_data)}")
            print(f"[DEBUG] Header row index: {new_header_row_index}")
            if new_header_row_index < len(sheet_data):
                print(f"[DEBUG] Header row content: {sheet_data[new_header_row_index]}")
            if len(sheet_data) > 0:
                print(f"[DEBUG] First row of sheet_data: {sheet_data[0]}")
            if len(sheet_data) > 1:
                print(f"[DEBUG] Second row of sheet_data: {sheet_data[1]}")
            
            # Refresh the UI for this sheet
            self._refresh_single_sheet_tab(sheet)
            
            # Update status
            self._update_status(f"Header row updated to row {new_header_row_index + 1} for sheet '{sheet.sheet_name}'")
            
            return True
            
        except Exception as e:
            logger.error(f"Error reprocessing sheet with header row {new_header_row_index}: {e}")
            messagebox.showerror("Error", f"Failed to reprocess sheet: {str(e)}")
            return False
    
    def _refresh_single_sheet_tab(self, sheet):
        """Refresh the UI for a single sheet tab"""
        try:
            # Find the sheet tab in the notebook
            if not hasattr(self, 'file_mapping') or not self.file_mapping:
                return
            
            # Get the sheet treeview
            tree = self.sheet_treeviews.get(sheet.sheet_name)
            if not tree:
                return
            
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
            
            # Required types (base required + successfully mapped new columns)
            if self.column_mapper and hasattr(self.column_mapper, 'config'):
                base_required_types = {col_type.value for col_type in self.column_mapper.config.get_required_columns()}
            else:
                base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
            
            # Add new columns as "required" if they are successfully mapped (confidence > 0)
            required_types = base_required_types.copy()
            if hasattr(sheet, 'column_mappings'):
                for mapping in sheet.column_mappings:
                    mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                    confidence = getattr(mapping, 'confidence', 0)
                    # Treat scope, manhours, wage as required if successfully mapped
                    if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                        required_types.add(mapped_type)
            
            # Repopulate treeview with updated column mappings
            if hasattr(sheet, 'column_mappings'):
                for mapping in sheet.column_mappings:
                    confidence = getattr(mapping, 'confidence', 0)
                    mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                    required = mapped_type in required_types
                    original_header = getattr(mapping, 'original_header', 'Unknown')
                    
                    # Determine if this mapping was user-edited
                    actions = "Manual Header" if getattr(mapping, 'user_edited', False) else "Auto-detected"
                    
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
            
        except Exception as e:
            logger.error(f"Error refreshing sheet tab: {e}")
            # Don't show error dialog as this is a background refresh

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
            "description", "quantity", "unit_price", "total_price", "unit", "code", "scope", "manhours", "wage", "ignore"
        ]
        # Base required types + new columns if successfully mapped
        base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        required_types = base_required_types.copy()
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                confidence = getattr(mapping, 'confidence', 0)
                if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                    required_types.add(mapped_type)
        # Dialog to edit mapped type
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Column: {column_name}")
        dialog.geometry("500x650")
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
        # Base required types + new columns if successfully mapped
        base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        required_types = base_required_types.copy()
        if hasattr(source_sheet, 'column_mappings'):
            for mapping in source_sheet.column_mappings:
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                confidence = getattr(mapping, 'confidence', 0)
                if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                    required_types.add(mapped_type)
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
            
        # CRITICAL VALIDATION: Check for missing required columns before proceeding
        validation_result = self._validate_required_columns()
        if not validation_result['is_valid']:
            self._show_missing_columns_warning(validation_result)
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
                
                # CRITICAL FIX: Check if header row was manually changed and adjust sheet data accordingly
                header_row_index = 0  # Default header row index
                if hasattr(self, 'controller') and self.controller:
                    # Find the file key by matching the file mapping
                    file_key = None
                    for key, file_data in self.controller.current_files.items():
                        if file_data.get('file_mapping') == self.file_mapping:
                            file_key = key
                            break
                    
                    if file_key and hasattr(self.controller, 'current_files') and file_key in self.controller.current_files:
                        processor_results = self.controller.current_files[file_key].get('processor_results', {})
                        header_row_indices = processor_results.get('header_row_indices', {})
                        
                        # Use the manually set header row index if available
                        if sheet.sheet_name in header_row_indices:
                            header_row_index = header_row_indices[sheet.sheet_name]
                            print(f"[DEBUG] Using manually set header row index {header_row_index} for sheet '{sheet.sheet_name}'")
                        elif hasattr(sheet, 'header_row_index'):
                            header_row_index = sheet.header_row_index
                
                # Remove the header row from sheet data before classification to prevent confusion
                # The row classifier should only classify data rows, not header rows
                if header_row_index < len(sheet_data):
                    # Create a copy of sheet data without the header row
                    data_rows = sheet_data[:header_row_index] + sheet_data[header_row_index + 1:]
                    print(f"[DEBUG] Removed header row {header_row_index} from sheet '{sheet.sheet_name}' data for classification")
                else:
                    data_rows = sheet_data
                
                # Convert column mappings to the format expected by row classifier
                # Use the current column mappings from the sheet (which reflect any manual changes)
                column_mapping_dict = {}
                
                # Use the sheet's current column mappings
                for col_mapping in sheet.column_mappings:
                    try:
                        # Convert string column type to ColumnType enum
                        col_type = ColumnType(col_mapping.mapped_type)
                        # Use the column index as-is (it's already 0-based from the mapping result)
                        column_mapping_dict[col_mapping.column_index] = col_type
                    except ValueError:
                        # Skip unknown column types
                        continue
                
                # DEBUG: Log the column mapping and first few rows to diagnose indexing
                # Only debug the "Miscellaneous" sheet where Bank Guarantee is located
                if sheet.sheet_name == "Miscellaneous":
                    # print(f"[DEBUG] Column mapping for {sheet.sheet_name}: {column_mapping_dict}")
                    if sheet_data:
                        # print(f"[DEBUG] First row data: {sheet_data[0] if len(sheet_data) > 0 else 'No data'}")
                        # print(f"[DEBUG] Row data length: {len(sheet_data[0]) if sheet_data else 0}")
                        
                        # Show headers to understand column mapping
                        if hasattr(sheet, 'column_mappings'):
                            # print(f"[DEBUG] Headers for {sheet.sheet_name}:")
                            for i, col_mapping in enumerate(sheet.column_mappings):
                                header = getattr(col_mapping, 'original_header', f'Column_{i}')
                                mapped_type = getattr(col_mapping, 'mapped_type', 'unknown')
                                confidence = getattr(col_mapping, 'confidence', 0)
                                # print(f"[DEBUG] Column {col_mapping.column_index}: '{header}' -> {mapped_type} (confidence: {confidence:.2f})")
                        
                        # Show first 10 rows to find Bank Guarantee
                        for i in range(min(10, len(sheet_data))):
                            if any('Bank Guarantee' in str(cell) for cell in sheet_data[i]):
                                # print(f"[DEBUG] FOUND BANK GUARANTEE at row {i}: {sheet_data[i]}")
                                pass
                            else:
                                # print(f"[DEBUG] Row {i}: {sheet_data[i]}")
                                pass
                
                # Perform row classification using the data rows (without header row)
                row_classification_result = row_classifier.classify_rows(data_rows, column_mapping_dict, sheet.sheet_name)
                
                # Update the sheet's row classifications
                sheet.row_classifications = []
                for row_class in row_classification_result.classifications:
                    from core.mapping_generator import RowClassificationInfo
                    # CRITICAL FIX: Adjust row index to account for removed header row
                    # The row classifier processed data WITHOUT the header row, so its indices
                    # are already shifted down. We need to shift them back up to match the original data.
                    adjusted_row_index = row_class.row_index
                    if row_class.row_index >= header_row_index:
                        adjusted_row_index = row_class.row_index + 1
                    
                    print(f"[DEBUG] Row classification: original_index={row_class.row_index}, header_row={header_row_index}, adjusted_index={adjusted_row_index}")
                    
                    row_info = RowClassificationInfo(
                        row_index=adjusted_row_index,
                        row_type=row_class.row_type.value,
                        confidence=row_class.confidence,
                        completeness_score=row_class.completeness_score,
                        hierarchical_level=row_class.hierarchical_level,
                        section_title=row_class.section_title,
                        validation_errors=row_class.validation_errors,
                        reasoning=row_class.reasoning,
                        position=row_class.position
                    )
                    sheet.row_classifications.append(row_info)
                
                # Update sheet statistics
                sheet.row_count = len(data_rows)  # Use data rows count (excluding header)
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
                pass
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

    def _validate_required_columns(self):
        """
        Validate that all required columns are properly mapped across all sheets.
        Returns a dictionary with validation results.
        """
        # Define required columns
        base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        
        validation_result = {
            'is_valid': True,
            'missing_columns': {},  # sheet_name -> [missing_column_types]
            'unmapped_sheets': [],
            'total_sheets': 0,
            'valid_sheets': 0
        }
        
        if not self.file_mapping or not hasattr(self.file_mapping, 'sheets'):
            validation_result['is_valid'] = False
            return validation_result
        
        for sheet in self.file_mapping.sheets:
            validation_result['total_sheets'] += 1
            sheet_name = sheet.sheet_name
            
            # Get mapped types for this sheet
            mapped_types = set()
            if hasattr(sheet, 'column_mappings'):
                for mapping in sheet.column_mappings:
                    mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                    confidence = getattr(mapping, 'confidence', 0)
                    # Only count columns with reasonable confidence or user-edited
                    if confidence > 0 or getattr(mapping, 'user_edited', False):
                        mapped_types.add(mapped_type)
            
            # Check for missing required columns
            missing_columns = []
            for required_type in base_required_types:
                if required_type not in mapped_types:
                    missing_columns.append(required_type)
            
            if missing_columns:
                validation_result['is_valid'] = False
                validation_result['missing_columns'][sheet_name] = missing_columns
            else:
                validation_result['valid_sheets'] += 1
        
        return validation_result
    
    def _show_missing_columns_warning(self, validation_result):
        """
        Show a detailed warning dialog about missing required columns.
        """
        from tkinter import messagebox
        
        missing_info = validation_result['missing_columns']
        total_sheets = validation_result['total_sheets']
        valid_sheets = validation_result['valid_sheets']
        
        # Build detailed message
        message_parts = [
            f"⚠️ COLUMN MAPPING INCOMPLETE ⚠️",
            f"",
            f"Cannot proceed to row mapping because some required columns are not mapped.",
            f"",
            f"Status: {valid_sheets}/{total_sheets} sheets have complete column mappings.",
            f""
        ]
        
        if missing_info:
            message_parts.append("Missing required columns by sheet:")
            message_parts.append("")
            
            for sheet_name, missing_columns in missing_info.items():
                message_parts.append(f"📋 {sheet_name}:")
                for col in missing_columns:
                    message_parts.append(f"   • {col}")
                message_parts.append("")
        
        message_parts.extend([
            "Required columns for BOQ processing:",
            "• description (item descriptions)",
            "• quantity (quantities)",  
            "• unit (units of measurement)",
            "• unit_price (unit prices)",
            "• total_price (total prices)",
            "• code (position codes)",
            "",
            "Please:",
            "1. Review the column mappings in each sheet tab",
            "2. Double-click columns to edit their mapping",
            "3. Ensure all required columns are properly mapped",
            "4. Try row mapping again"
        ])
        
        full_message = "\n".join(message_parts)
        
        messagebox.showerror(
            "Missing Required Columns",
            full_message
        )
        
        self._update_status(f"Column mapping incomplete: {total_sheets - valid_sheets} sheets need attention")

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
        # Base required types + new columns if successfully mapped
        base_required_types = {"description", "quantity", "unit_price", "total_price", "unit", "code"}
        required_types = base_required_types.copy()
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                confidence = getattr(mapping, 'confidence', 0)
                if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                    required_types.add(mapped_type)
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
                base_required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
            else:
                base_required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
            
            # Add new columns as "required" if they are successfully mapped (confidence > 0)
            required_types = base_required_types.copy()
            if hasattr(sheet, 'column_mappings'):
                for mapping in sheet.column_mappings:
                    mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                    confidence = getattr(mapping, 'confidence', 0)
                    # Treat scope, manhours, wage as required if successfully mapped
                    if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                        if mapped_type not in required_types:
                            required_types.append(mapped_type)
            
            # Define the correct column order for display
            display_column_order = ["code", "description", "unit", "quantity", "unit_price", "total_price", "scope", "manhours", "wage"]
            
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
                    tree.column(col, width=120 if col != "#" else 40, anchor=tk.W, minwidth=50, stretch=False)
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
                        if col in ['unit_price', 'total_price', 'wage']:
                            val = format_number(val, is_currency=True)
                        elif col in ['quantity', 'manhours']:
                            val = format_number(val, is_currency=False)
                            # Special formatting for manhours - only 2 decimals
                            if col == 'manhours' and val and val != '':
                                try:
                                    num_val = float(str(val).replace(',', '.'))
                                    val = f"{num_val:.2f}"
                                except:
                                    pass
                        
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
                # DEBUG: Log sheet data structure
                print(f"[DEBUG] Processing sheet: {sheet.sheet_name}")
                if hasattr(sheet, 'sheet_data'):
                    print(f"[DEBUG] Sheet data length: {len(sheet.sheet_data)}")
                    if len(sheet.sheet_data) > 0:
                        print(f"[DEBUG] First row of sheet data: {sheet.sheet_data[0]}")
                    if len(sheet.sheet_data) > 1:
                        print(f"[DEBUG] Second row of sheet data: {sheet.sheet_data[1]}")
                else:
                    print(f"[DEBUG] No sheet_data attribute found for {sheet.sheet_name}")
                
                # DEBUG: Log column mappings
                print(f"[DEBUG] Column mappings for {sheet.sheet_name}:")
                for i, cm in enumerate(sheet.column_mappings):
                    print(f"[DEBUG]   {i}: {cm.column_index} -> {cm.original_header} -> {cm.mapped_type}")
                
                # DEBUG: Log col_headers array
                print(f"[DEBUG] col_headers array: {col_headers}")
                
                # For each row classification, get the row data
                for rc in getattr(sheet, 'row_classifications', []):
                    # Only include valid rows
                    if not validity_dict.get(rc.row_index, True):
                        continue
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                            print(f"[DEBUG] Row {rc.row_index}: extracted data = {row_data}")
                        except Exception as e:
                            print(f"[DEBUG] Row {rc.row_index}: extraction failed = {e}")
                            row_data = None
                    if row_data is None:
                        row_data = []
                        print(f"[DEBUG] Row {rc.row_index}: using empty data")
                    
                    # Build dict for DataFrame
                    row_dict = {}
                    for cm in sheet.column_mappings:
                        mapped_type = getattr(cm, 'mapped_type', None)
                        if not mapped_type:
                            continue
                        idx = cm.column_index
                        row_dict[mapped_type] = row_data[idx] if idx < len(row_data) else ''
                    # Ensure all expected columns are present
                    for mt in col_headers:
                        if mt not in row_dict:
                            row_dict[mt] = ''
                    
                    # DEBUG: Log specific column values and check for data swapping
                    if 'description' in row_dict:
                        desc_val = row_dict['description']
                        print(f"[DEBUG] Row {rc.row_index}: description = '{desc_val}'")
                        # Check if description looks like a code
                        if isinstance(desc_val, str) and len(desc_val) < 30 and ('.' in desc_val or '/' in desc_val):
                            print(f"[DEBUG] ⚠️  Row {rc.row_index}: DESCRIPTION LOOKS LIKE CODE: '{desc_val}'")
                    if 'code' in row_dict:
                        code_val = row_dict['code']
                        print(f"[DEBUG] Row {rc.row_index}: code = '{code_val}'")
                        # Check if code looks like a description
                        if isinstance(code_val, str) and len(code_val) > 30:
                            print(f"[DEBUG] ⚠️  Row {rc.row_index}: CODE LOOKS LIKE DESCRIPTION: '{code_val[:50]}...'")
                    
                    # DEBUG: Log the full row_dict to see all values
                    print(f"[DEBUG] Row {rc.row_index}: Full row_dict = {row_dict}")
                    
                    row_dict['Source_Sheet'] = sheet.sheet_name
                    # Preserve position information for future row matching
                    row_dict['Position'] = getattr(rc, 'position', None)
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
                
                # DEBUG: Additional logging for categorization DataFrame
                print(f"[DEBUG] CATEGORIZATION DataFrame created with {len(df)} rows")
                print(f"[DEBUG] CATEGORIZATION DataFrame columns: {list(df.columns)}")
                if 'Description' in df.columns:
                    print(f"[DEBUG] CATEGORIZATION First 5 Description values: {df['Description'].head(5).tolist()}")
                    # Check if descriptions look like codes
                    desc_sample = df['Description'].head(10).tolist()
                    code_like_count = sum(1 for desc in desc_sample if isinstance(desc, str) and len(desc) < 20 and '.' in desc)
                    print(f"[DEBUG] CATEGORIZATION {code_like_count}/{len(desc_sample)} descriptions look like codes")
                if 'code' in df.columns:
                    print(f"[DEBUG] CATEGORIZATION First 5 code values: {df['code'].head(5).tolist()}")
                print(f"[DEBUG] CATEGORIZATION First 3 complete rows:")
                print(df.head(3).to_string())
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
            # print("[DEBUG] Current tab path:", current_tab_path)
            # print("[DEBUG] Current tab widget:", current_tab)
            # print("[DEBUG] File mapping tabs:", [str(file_data['file_mapping'].tab) for file_data in self.controller.current_files.values()])
            # print("[DEBUG] Number of files:", len(self.controller.current_files))
            
            for file_key, file_data in self.controller.current_files.items():
                # print("[DEBUG] Checking file_key:", file_key)
                # print("[DEBUG] file_data['file_mapping'].tab:", file_data['file_mapping'].tab)
                # print("[DEBUG] hasattr check:", hasattr(file_data['file_mapping'], 'tab'))
                # print("[DEBUG] tab comparison:", str(file_data['file_mapping'].tab) == str(current_tab_path))
                
                if hasattr(file_data['file_mapping'], 'tab') and str(file_data['file_mapping'].tab) == str(current_tab_path):
                    # print("[DEBUG] Found matching tab, storing data...")
                    # Store the categorized data
                    file_data['categorized_dataframe'] = final_dataframe
                    file_data['categorization_result'] = categorization_result
                    # Update the file mapping
                    file_mapping = file_data['file_mapping']
                    file_mapping.categorized_dataframe = final_dataframe
                    file_mapping.categorization_result = categorization_result
                    # print("[DEBUG] About to call _show_final_categorized_data...")
                    # Show the final data grid in the main window - use the actual tab widget from file_mapping
                    self._show_final_categorized_data(file_mapping.tab, final_dataframe, categorization_result)
                    # Refresh the new summary grid after categorization using centralized method
                    print(f"[DEBUG] Calling centralized summary grid refresh after categorization")
                    self._refresh_summary_grid_centralized()
                    self._update_status("Categorization completed successfully - showing final data")
                    # print("[DEBUG] _show_final_categorized_data call completed")
                    break
                else:
                    # print("[DEBUG] Tab mismatch or no tab attribute")
                    pass
        except Exception as e:
            # print("[DEBUG] Exception in _on_categorization_complete:", e)
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
            std_order = ['code', 'sheet', 'category', 'description', 'quantity', 'unit_price', 'total_price', 'unit', 'scope', 'manhours', 'wage']
            for t in std_order:
                if t not in ordered_types:
                    ordered_types.append(t)
            return ordered_types
        else:
            return ['code', 'sheet', 'category', 'description', 'quantity', 'unit_price', 'total_price', 'unit', 'scope', 'manhours', 'wage']

    def _build_final_grid_dataframe(self, file_mapping):
        import pandas as pd
        rows = []
        # Determine required types and display order as in row review
        if self.column_mapper and hasattr(self.column_mapper, 'config'):
            base_required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
        else:
            base_required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
        
        # Add new columns as "required" if they are successfully mapped
        required_types = base_required_types.copy()
        for sheet in getattr(file_mapping, 'sheets', []):
            if hasattr(sheet, 'column_mappings'):
                for mapping in sheet.column_mappings:
                    mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                    confidence = getattr(mapping, 'confidence', 0)
                    if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                        if mapped_type not in required_types:
                            required_types.append(mapped_type)
        # Row review display order
        display_column_order = ["code", "sheet", "category", "description", "unit", "quantity", "unit_price", "total_price", "scope", "manhours", "wage"]
        
        # Helper function to parse numbers - preserve empty values for manhours and wage
        def parse_number(val, preserve_empty=False):
            if isinstance(val, (int, float)):
                return float(val)
            if pd.isna(val) or val == '' or (isinstance(val, str) and val.strip() == ''):
                return '' if preserve_empty else 0.0
            s = str(val).replace('\u202f', '').replace(' ', '').replace(',', '.')
            try:
                return float(s)
            except Exception:
                return '' if preserve_empty else 0.0
        
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
                            # Parse numeric values - preserve empty values for manhours and wage
                            if col in ['quantity', 'unit_price', 'total_price']:
                                val = parse_number(val)
                            elif col in ['manhours', 'wage']:
                                val = parse_number(val, preserve_empty=True)
                            row_dict[col] = val
                    
                    # Preserve position information for future row matching (not displayed but stored)
                    row_dict['position'] = getattr(rc, 'position', None)
                    
                    rows.append(row_dict)
        # Build DataFrame
        if rows:
            df = pd.DataFrame(rows)
            # Ensure numeric columns are properly typed
            for col in ['quantity', 'unit_price', 'total_price']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    df[col] = df[col].fillna(0)
            # Handle manhours and wage separately - preserve empty values
            for col in ['manhours', 'wage']:
                if col in df.columns:
                    # Convert to numeric but keep empty strings as empty strings
                    df[col] = df[col].apply(lambda x: pd.to_numeric(x, errors='coerce') if x != '' else '')
                    # Only fill NaN with 0 if it was originally a number that failed to parse
                    df[col] = df[col].apply(lambda x: 0 if pd.isna(x) and x != '' else x)
            # Only keep columns in display order that are present
            columns = [col for col in display_column_order if col in df.columns]
            df = df[columns]
        else:
            df = pd.DataFrame(data=[], columns=display_column_order)
        return df

    def _show_final_categorized_data(self, tab, final_dataframe, categorization_result):
        # print("[DEBUG] _show_final_categorized_data called for tab:", tab)
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
            final_display_columns = list(display_df.columns)
            # print(f"[DEBUG] Using loaded DataFrame directly with shape: {display_df.shape}")
            
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
            main_frame.grid_rowconfigure(0, weight=0)  # Title
            main_frame.grid_rowconfigure(1, weight=0)  # Instructions
            main_frame.grid_rowconfigure(2, weight=0)  # NEW Summary Grid (Supplier, Project Name, Date, Total Price)
            main_frame.grid_rowconfigure(3, weight=1)  # Data grid (expandable)
            main_frame.grid_rowconfigure(4, weight=0)  # EXISTING Summary frame
            main_frame.grid_rowconfigure(5, weight=0)  # Button frame (always visible)
            main_frame.grid_columnconfigure(0, weight=1)
            # Title and instructions
            title_label = ttk.Label(main_frame, text="Final Categorized Data", font=("TkDefaultFont", 14, "bold"))
            title_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
            instructions = """
            Review the categorized data below. You can make corrections by double-clicking the category cell and selecting a value from the dropdown.\nChanges will be saved when you click 'Apply Changes', 'Summarize', 'Save Analysis', or 'Export Data'.
            """
            instruction_label = ttk.Label(main_frame, text=instructions, wraplength=800, justify=tk.LEFT)
            instruction_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
            
            # --- SCROLLABLE MAIN CONTENT AREA ---
            tab.grid_rowconfigure(1, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            canvas = tk.Canvas(tab)
            canvas.grid(row=1, column=0, sticky=tk.NSEW)
            vscroll = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
            vscroll.grid(row=1, column=1, sticky=tk.NS)
            canvas.configure(yscrollcommand=vscroll.set)
            # Create a frame inside the canvas
            scrollable_frame = ttk.Frame(canvas)
            scrollable_frame.bind(
                "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            # Ensure the scrollable frame always matches the canvas width
            def _on_canvas_configure(event):
                canvas.itemconfig('main_frame', width=event.width)
            canvas.bind('<Configure>', _on_canvas_configure)
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags='main_frame')
            # Make the scrollable frame expand
            scrollable_frame.grid_rowconfigure(0, weight=1)
            scrollable_frame.grid_columnconfigure(0, weight=1)
            main_frame = scrollable_frame
            main_frame.grid_columnconfigure(0, weight=1)
            # --- MAIN DATASET GRID (Final Data) ---
            MAX_TREE_HEIGHT = 20
            DATASET_FRAME_HEIGHT = 400
            tree_frame = ttk.Frame(main_frame, height=DATASET_FRAME_HEIGHT)
            tree_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
            tree_frame.grid_propagate(False)
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)
            tree = ttk.Treeview(tree_frame, columns=final_display_columns, show='headings', height=MAX_TREE_HEIGHT)
            for col in final_display_columns:
                tree.heading(col, text=col.capitalize() if col != '#' else '#')
                if col == 'description':
                    width = 250
                else:
                    width = self._calculate_column_width(display_df, col, col.capitalize())
                tree.column(col, width=width, minwidth=50, stretch=True)
            tree.grid(row=0, column=0, sticky=tk.NSEW)
            dataset_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=dataset_vsb.set)
            dataset_vsb.grid(row=0, column=1, sticky=tk.NS)
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=hsb.set)
            hsb.grid(row=1, column=0, sticky=tk.EW)
            style = ttk.Style(tree)
            style.map('Treeview', background=[('selected', '#B3E5FC')], foreground=[('selected', 'black')])
            self._populate_final_data_treeview(tree, display_df, final_display_columns)
            self._enable_final_data_editing(tree, display_df)

            # --- GLOBAL SUMMARY GRID (BOQ Summary Overview) ---
            if not hasattr(self, 'global_summary_frame') or self.global_summary_frame is None:
                self.global_summary_frame = ttk.LabelFrame(main_frame, text="BOQ Summary Overview")
                self.global_summary_frame.grid(row=3, column=0, sticky=tk.EW, pady=(0, 10))
                self.global_summary_frame.grid_columnconfigure(0, weight=1)
                print(f"[DEBUG] Created new global summary frame")
            else:
                self.global_summary_frame.grid(row=3, column=0, sticky=tk.EW, pady=(0, 10))
                print(f"[DEBUG] Reusing existing global summary frame")
            self._create_new_summary_grid(self.global_summary_frame, tab)

            # --- DETAILED SUMMARY GRID (Summarize) ---
            summary_frame = ttk.Frame(main_frame)
            summary_frame.grid(row=4, column=0, sticky=tk.EW, pady=(0, 10))
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
                        # print(f"[DEBUG] Failed to parse number '{val}': {e}")
                        return 0.0
                
                # Get the current DataFrame (now uses pretty categories directly)
                df = tab.final_dataframe if hasattr(tab, 'final_dataframe') else display_df
                # print("[DEBUG] DataFrame columns:", df.columns.tolist())
                # print("[DEBUG] First few rows of DataFrame:")
                print(df.head().to_string())
                
                # Check if this is a comparison dataset
                is_comparison = self._is_comparison_dataset(df)
                # print(f"[DEBUG] Is comparison dataset: {is_comparison}")
                
                # Remove old summary tree if present
                for widget in summary_frame.winfo_children():
                    widget.destroy()
                
                # Configure the grid for the summary frame to allow the scrollbar to show
                summary_frame.grid_columnconfigure(0, weight=1)
                
                if is_comparison:
                    # Handle comparison dataset - create separate rows for each offer
                    # print("[DEBUG] Creating comparison summary")
                    
                    # Find all offer columns
                    offer_columns = {}
                    for col in df.columns:
                        if col.startswith(('total_price[', 'total_price_')):
                            if '[' in col and ']' in col:
                                offer_name = col.split('[')[1].split(']')[0]
                            elif '_' in col:
                                offer_name = col.split('_', 1)[1]
                            else:
                                continue
                            offer_columns[offer_name] = col
                    
                    # print(f"[DEBUG] Found offer columns: {offer_columns}")
                    
                    if not offer_columns:
                        summary_frame.grid_remove()
                        return
                    
                    # Create summary columns: Offer + categories
                    summary_columns = ['Offer'] + categories_pretty
                    
                    # Calculate height based on number of offers
                    tree_height = max(2, len(offer_columns))
                    summary_tree = ttk.Treeview(summary_frame, columns=summary_columns, show='headings', height=tree_height)
                    
                    for col in summary_columns:
                        summary_tree.heading(col, text=col)
                        summary_tree.column(col, width=150, minwidth=100)
                    
                    # Create a row for each offer
                    for offer_name, price_col in offer_columns.items():
                        # Parse the price column
                        df[price_col] = df[price_col].apply(parse_number)
                        
                        # Group by category for this offer
                        if 'category' in df.columns:
                            summary_dict = df.groupby('category')[price_col].sum().to_dict()
                            
                            # Ensure all categories are present
                            final_summary = {}
                            for cat_pretty in categories_pretty:
                                final_summary[cat_pretty] = summary_dict.get(cat_pretty, 0.0)
                        else:
                            final_summary = {cat: 0.0 for cat in categories_pretty}
                        
                        # Create display values for this offer
                        summary_values = [offer_name] + [final_summary[cat] for cat in categories_pretty]
                        
                        # Format values for display
                        display_values = []
                        for i, val in enumerate(summary_values):
                            if i == 0:  # Offer label
                                display_values.append(str(val))
                            else:  # Numeric values
                                display_values.append(format_number_eu(val))
                        
                        summary_tree.insert('', 'end', values=display_values, tags=('offer',))
                        # print(f"[DEBUG] Added summary row for {offer_name}: {display_values[:3]}...")
                    
                    # Add horizontal scrollbar
                    hsb_summary = ttk.Scrollbar(summary_frame, orient=tk.HORIZONTAL, command=summary_tree.xview)
                    summary_tree.configure(xscrollcommand=hsb_summary.set)
                    summary_tree.grid(row=0, column=0, sticky=tk.EW)
                    hsb_summary.grid(row=1, column=0, sticky=tk.EW)
                    
                else:
                    # Handle single offer dataset (original logic)
                    # print("[DEBUG] Creating single offer summary")
                    
                    # Find the correct total price column
                    price_col = None
                    possible_price_cols = ['total_price', 'Total_price']
                    
                    for col in possible_price_cols:
                        if col in df.columns:
                            price_col = col
                            break
                    
                    # print(f"[DEBUG] Using price column: {price_col}")
                    if price_col and price_col in df.columns:
                        df[price_col] = df[price_col].apply(parse_number)
                        # print(f"[DEBUG] Price column after parsing:")
                        print(df[price_col].head().to_string())
                    else:
                        # print("[DEBUG] No valid price column found for summary")
                        summary_frame.grid_remove()
                        return
                    
                    # Group by category
                    cat_col = 'category'
                    if cat_col in df.columns and price_col and price_col in df.columns:
                        # Categories are already in pretty format, just group directly
                        # print(f"[DEBUG] Unique categories before grouping:", df[cat_col].unique())
                        
                        # Create the summary dictionary with pretty category names
                        summary_dict = df.groupby(cat_col)[price_col].sum().to_dict()
                        # print("[DEBUG] Summary dict after grouping:", summary_dict)
                        
                        # No need for complex mapping - categories are already pretty
                        # Just ensure we have all predefined categories with zero values if not present
                        final_summary = {}
                        for cat_pretty in categories_pretty:
                            final_summary[cat_pretty] = summary_dict.get(cat_pretty, 0.0)
                        
                        # print("[DEBUG] Final summary:", final_summary)
                    else:
                        # print("[DEBUG] No category column or price column found!")
                        final_summary = {cat: 0.0 for cat in categories_pretty}
                    
                    offer_label = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else 'Offer'
                    summary_columns = ['Offer'] + categories_pretty
                    
                    # Use the final summary to get values in the correct order
                    summary_values = [offer_label] + [final_summary[cat] for cat in categories_pretty]
                    
                    # print("[DEBUG] Final summary values:", summary_values)
                    
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
                            display_values.append(format_number_eu(val))
                    
                    summary_tree.insert('', 'end', values=display_values, tags=('offer',))
                
                # Add horizontal scrollbar for the summary grid
                hsb_summary = ttk.Scrollbar(summary_frame, orient=tk.HORIZONTAL, command=summary_tree.xview)
                summary_tree.configure(xscrollcommand=hsb_summary.set)
                summary_tree.grid(row=0, column=0, sticky=tk.EW)
                hsb_summary.grid(row=1, column=0, sticky=tk.EW)
                
                # Add copy button for the existing summary grid
                copy_frame = ttk.Frame(summary_frame)
                copy_frame.grid(row=2, column=0, sticky=tk.E, pady=(5, 0))
                
                copy_button = ttk.Button(copy_frame, text="📋 Copy to Clipboard", 
                                       command=lambda: self._copy_grid_to_clipboard(summary_tree, "Detailed Summary"))
                copy_button.pack(side=tk.RIGHT)
                
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
                    
                    # Get the current full DataFrame - always use tab.final_dataframe as primary source
                    current_df = getattr(tab, 'final_dataframe', None)
                    if current_df is None:
                        # Fallback to display_df only if tab.final_dataframe doesn't exist
                        current_df = display_df
                    
                    # Get the current display columns - dynamically from the actual DataFrame
                    current_display_columns = list(current_df.columns)
                    
                    # Toggle filter: if already filtered to this, remove; else filter
                    if getattr(tab, '_active_category_filter', None) == category_pretty:
                        tab._active_category_filter = None
                        filtered_df = current_df
                    else:
                        tab._active_category_filter = category_pretty
                        # Filter by pretty category directly - no conversion needed
                        if 'category' in current_df.columns:
                            filtered_df = current_df[current_df['category'] == category_pretty]
                        else:
                            # If no category column, show empty result
                            filtered_df = current_df.iloc[0:0]  # Empty DataFrame with same structure
                    
                    # Repopulate the main grid with the filtered DataFrame using current columns
                    self._populate_final_data_treeview(tab.final_data_tree, filtered_df, current_display_columns)
                    # Update the reference so further edits work on the filtered view
                    tab._filtered_dataframe = filtered_df
                
                summary_tree.bind('<Double-1>', on_summary_double_click)
                # --- END CATEGORY FILTERING FEATURE ---
            # --- END SUMMARY GRID PLACEHOLDER ---
            # Button frame at the bottom - always visible
            button_frame = ttk.Frame(main_frame)
            button_frame.grid(row=5, column=0, sticky=tk.EW, pady=(10, 0))
            button_frame.grid_columnconfigure(0, weight=1)  # Allow buttons to expand
            
            # Create a centered button container
            button_container = ttk.Frame(button_frame)
            button_container.grid(row=0, column=0)
            
            summarize_button = ttk.Button(button_container, text="Summarize", command=show_summary_grid)
            summarize_button.pack(side=tk.LEFT, padx=(0, 5))
            save_analysis_button = ttk.Button(button_container, text="Save Analysis", command=lambda: self._save_analysis(tab))
            save_analysis_button.pack(side=tk.LEFT, padx=(0, 5))
            save_mappings_button = ttk.Button(button_container, text="Save Mappings", command=lambda: self._save_mappings(tab))
            save_mappings_button.pack(side=tk.LEFT, padx=(0, 5))
            compare_full_button = ttk.Button(button_container, text="Compare Full", command=lambda: self._compare_full(tab))
            compare_full_button.pack(side=tk.LEFT, padx=(0, 5))
            export_button = ttk.Button(button_container, text="Export Data", 
                                      command=lambda: self._export_final_data(tab.final_dataframe, tab))
            export_button.pack(side=tk.LEFT, padx=(0, 5))
            
            # Store button references for state management
            tab.compare_full_button = compare_full_button
            tab.final_data_tree = tree
            tab.final_dataframe = display_df
            tab.categorization_result = categorization_result
        except Exception as e:
            # print("[DEBUG] Exception in _show_final_categorized_data:", e)
            import traceback
            traceback.print_exc()
    
    def _populate_final_data_treeview(self, tree, dataframe, columns):
        """Populate the treeview with data from the final DataFrame, using pretty categories directly."""
        # print(f"[DEBUG] _populate_final_data_treeview called with DataFrame shape: {dataframe.shape}")
        # print(f"[DEBUG] DataFrame columns: {dataframe.columns.tolist()}")
        # print(f"[DEBUG] Requested columns: {columns}")
        
        # Unit column validation - ensure it's present in the data
        if 'unit' in columns and 'unit' not in dataframe.columns:
            print(f"[WARNING] Unit column requested but not found in DataFrame columns: {dataframe.columns.tolist()}")
        
        # Helper to format numbers consistently
        def format_number(val, is_currency=False):
            try:
                if pd.isna(val):
                    return ""
                if isinstance(val, str):
                    # Remove any existing formatting
                    val = val.replace(' ', '').replace('\u202f', '').replace(',', '.')
                num = float(val)
                return format_number_eu(num)
            except (ValueError, TypeError):
                return str(val)
        
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # print(f"[DEBUG] Cleared existing items, now adding {len(dataframe)} rows")
        
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
                elif col == 'unit':
                    # Unit column handling
                    unit_value = str(value) if value != '' else ''
                    values.append(unit_value)
                elif col in ['quantity', 'unit_price', 'total_price', 'wage'] or col.startswith(('quantity[', 'unit_price[', 'total_price[', 'wage[', 'quantity_', 'unit_price_', 'total_price_', 'wage_')):
                    # Format numeric columns (including comparison columns)
                    values.append(format_number(value))
                elif col in ['manhours'] or col.startswith(('manhours[', 'manhours_')):
                    # Special formatting for manhours - only 2 decimals
                    try:
                        if pd.isna(value) or value == '':
                            formatted_val = ""
                        else:
                            num_val = float(str(value).replace(',', '.'))
                            formatted_val = f"{num_val:.2f}".replace('.', ',')
                        values.append(formatted_val)
                    except:
                        values.append(str(value))
                else:
                    values.append(str(value))
            
            tree.insert('', 'end', values=values, tags=(f'row_{index}',))
        
        # print(f"[DEBUG] Finished populating treeview with {len(dataframe)} rows")
    
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
                            # print(f"[DEBUG] Error saving combo: {e}")
                            combo.destroy()
                    def cancel_combo(event=None):
                        combo.destroy()
                    combo.bind('<Return>', save_combo)
                    combo.bind('<FocusOut>', save_combo)
                    combo.bind('<Escape>', cancel_combo)
                    combo.focus()
            except Exception as e:
                # print(f"[DEBUG] Error in double-click editing: {e}")
                pass
        tree.bind('<Double-1>', on_double_click)
        # print("[DEBUG] Double-click binding added to treeview (category only)")
    
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
        """Export the final categorized data to a nicely formatted Excel file with dynamic formulas in summary."""
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
            # print(f"[DEBUG] Exporting with categories: {df['category'].unique() if 'category' in df.columns else 'No category column'}")
            
            # Ensure numeric columns are numbers (including comparison columns)
            for col in df.columns:
                if col in ['quantity', 'unit_price', 'total_price', 'manhours', 'wage'] or col.startswith(('quantity[', 'unit_price[', 'total_price[', 'manhours[', 'wage[', 'quantity_', 'unit_price_', 'total_price_', 'manhours_', 'wage_')):
                    df[col] = df[col].apply(parse_number)
            
            # Write to Excel with formatting
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Main Dataset', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Main Dataset']
                
                # Format headers
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)
                
                # Format numeric columns (including comparison columns) - 2 decimal places as requested
                num_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
                manhours_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})  # 2 decimals for manhours
                # Note: Column formatting is applied during autofit below to preserve both width and format
                
                # Data validation for category column
                if 'category' in df.columns:
                    cat_col_idx = df.columns.get_loc('category')
                    worksheet.data_validation(1, cat_col_idx, len(df), cat_col_idx, {
                        'validate': 'list',
                        'source': categories_pretty,
                        'input_message': 'Select a category from the list.'
                    })
                
                # Autofit columns while preserving number formatting
                for i, col in enumerate(df.columns):
                    maxlen = max(
                        [len(str(x)) for x in df[col].astype(str).values] + [len(str(col))]
                    )
                    width = min(maxlen + 2, 30)
                    
                    # Preserve number formatting when setting column width
                    if col in ['quantity', 'unit_price', 'total_price', 'wage'] or col.startswith(('quantity[', 'unit_price[', 'total_price[', 'wage[', 'quantity_', 'unit_price_', 'total_price_', 'wage_')):
                        worksheet.set_column(i, i, width, num_format)
                    elif col in ['manhours'] or col.startswith(('manhours[', 'manhours_')):
                        worksheet.set_column(i, i, width, manhours_format)
                    else:
                        worksheet.set_column(i, i, width)
                
                # Add summary sheet with formulas
                if tab and hasattr(tab, 'summary_frame') and hasattr(tab, 'final_dataframe'):
                    df_summary = tab.final_dataframe.copy()
                    
                    # Check if this is a comparison dataset
                    is_comparison = self._is_comparison_dataset(df_summary)
                    
                    # Create summary sheet
                    summary_ws = workbook.add_worksheet('Summary')
                    
                    if is_comparison:
                        # Handle comparison dataset - create separate rows for each offer with formulas
                        # print("[DEBUG] Exporting comparison summary with formulas")
                        
                        # Find all offer columns
                        offer_columns = {}
                        for col in df_summary.columns:
                            if col.startswith(('total_price[', 'total_price_')):
                                if '[' in col and ']' in col:
                                    offer_name = col.split('[')[1].split(']')[0]
                                elif '_' in col:
                                    offer_name = col.split('_', 1)[1]
                                else:
                                    continue
                                offer_columns[offer_name] = col
                        
                        # print(f"[DEBUG] Export found offer columns: {offer_columns}")
                        
                        if offer_columns:
                            # Write headers with formatting
                            summary_columns = ['Offer'] + categories_pretty
                            for col_idx, col in enumerate(summary_columns):
                                summary_ws.write(0, col_idx, col, header_format)
                            
                            # Create a row for each offer with formulas
                            row_idx = 1
                            for offer_name, price_col in offer_columns.items():
                                # Write offer name
                                summary_ws.write(row_idx, 0, offer_name)
                                
                                # Find the price column letter in the main dataset
                                price_col_idx = df.columns.get_loc(price_col)
                                price_col_letter = chr(65 + price_col_idx)  # Convert to Excel column letter
                                
                                # Find category column letter in the main dataset
                                try:
                                    cat_col_idx = df.columns.get_loc('category')
                                    cat_col_letter = chr(65 + cat_col_idx)
                                except KeyError:
                                    # Try alternative category column names
                                    category_col_names = ['Category', 'CATEGORY', 'category', 'cat', 'Cat']
                                    cat_col_letter = None
                                    for cat_col_name in category_col_names:
                                        try:
                                            cat_col_idx = df.columns.get_loc(cat_col_name)
                                            cat_col_letter = chr(65 + cat_col_idx)
                                            break
                                        except KeyError:
                                            continue
                                    
                                    if cat_col_letter is None:
                                        # No category column found, use static values
                                        for col_idx, category in enumerate(categories_pretty, 1):
                                            summary_ws.write_number(row_idx, col_idx, 0, num_format)
                                        continue
                                
                                # Write formulas for each category
                                for col_idx, category in enumerate(categories_pretty, 1):
                                    # Formula: =SUMIF('Main Dataset'!$category_col:$category_col, "category_name", 'Main Dataset'!$price_col:$price_col)
                                    try:
                                        formula = f"=SUMIF('Main Dataset'!${cat_col_letter}:${cat_col_letter}, \"{category}\", 'Main Dataset'!${price_col_letter}:${price_col_letter})"
                                        summary_ws.write_formula(row_idx, col_idx, formula, num_format)
                                    except NameError:
                                        # cat_col_letter not defined (no category column)
                                        summary_ws.write_number(row_idx, col_idx, 0, num_format)
                                
                                row_idx += 1
                                # print(f"[DEBUG] Exported summary row with formulas for {offer_name}")
                            
                            
                            
                            # Format columns
                            for i in range(len(summary_columns)):
                                if i == 0:
                                    summary_ws.set_column(i, i, 25)  # Offer column: text
                                else:
                                    summary_ws.set_column(i, i, 18, num_format)  # Category columns: number format
                            

                    
                    else:
                        # Handle single offer dataset with formulas
                        # print("[DEBUG] Exporting single offer summary with formulas")
                        
                        # Find the correct total price column
                        price_col = None
                        possible_price_cols = ['total_price', 'Total_price']
                        
                        for col in possible_price_cols:
                            if col in df_summary.columns:
                                price_col = col
                                break
                        
                        if price_col and price_col in df_summary.columns:
                            # Write headers with formatting
                            summary_columns = ['Offer'] + categories_pretty
                            for col_idx, col in enumerate(summary_columns):
                                summary_ws.write(0, col_idx, col, header_format)
                            
                            # Write offer name
                            offer_label = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else 'Offer'
                            summary_ws.write(1, 0, offer_label)
                            
                            # Find the price column letter in the main dataset
                            price_col_idx = df.columns.get_loc(price_col)
                            price_col_letter = chr(65 + price_col_idx)
                            
                            # Find category column letter in the main dataset
                            try:
                                cat_col_idx = df.columns.get_loc('category')
                                cat_col_letter = chr(65 + cat_col_idx)
                            except KeyError:
                                # Try alternative category column names
                                category_col_names = ['Category', 'CATEGORY', 'category', 'cat', 'Cat']
                                cat_col_letter = None
                                for cat_col_name in category_col_names:
                                    try:
                                        cat_col_idx = df.columns.get_loc(cat_col_name)
                                        cat_col_letter = chr(65 + cat_col_idx)
                                        break
                                    except KeyError:
                                        continue
                                
                                if cat_col_letter is None:
                                    # No category column found, use static values
                                    for col_idx, category in enumerate(categories_pretty, 1):
                                        summary_ws.write_number(1, col_idx, 0, num_format)
                                    return
                            
                            # Write formulas for each category
                            for col_idx, category in enumerate(categories_pretty, 1):
                                # Formula: =SUMIF('Main Dataset'!$category_col:$category_col, "category_name", 'Main Dataset'!$price_col:$price_col)
                                try:
                                    formula = f"=SUMIF('Main Dataset'!${cat_col_letter}:${cat_col_letter}, \"{category}\", 'Main Dataset'!${price_col_letter}:${price_col_letter})"
                                    summary_ws.write_formula(1, col_idx, formula, num_format)
                                except NameError:
                                    # cat_col_letter not defined (no category column)
                                    summary_ws.write_number(1, col_idx, 0, num_format)
                            
                            
                            # Format columns
                            for i in range(len(summary_columns)):
                                if i == 0:
                                    summary_ws.set_column(i, i, 25)  # Offer column: text
                                else:
                                    summary_ws.set_column(i, i, 18, num_format)  # Category columns: number format
                        else:
                            # No price column found, create empty summary with headers
                            summary_columns = ['Offer'] + categories_pretty
                            for col_idx, col in enumerate(summary_columns):
                                summary_ws.write(0, col_idx, col, header_format)
                            
                            offer_label = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else 'Offer'
                            summary_ws.write(1, 0, offer_label)
                            
                            # Write 0 for each category (no formulas possible)
                            for col_idx in range(1, len(summary_columns)):
                                summary_ws.write_number(1, col_idx, 0, num_format)
                            
                            # Format columns
                            for i in range(len(summary_columns)):
                                if i == 0:
                                    summary_ws.set_column(i, i, 25)  # Offer column: text
                                else:
                                    summary_ws.set_column(i, i, 18, num_format)  # Category columns: number format
                
                messagebox.showinfo("Success", f"Data exported to: {file_path}\n\nSummary sheet now contains dynamic formulas that automatically update when you modify the main dataset.")
                self._update_status(f"Exported data with formulas to: {file_path}")
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def run(self):
        """Start the main application loop."""
        self.root.mainloop()

    def _prompt_offer_info(self, is_first_boq: bool = True):
        """
        Prompt for comprehensive offer information
        
        Args:
            is_first_boq: True if this is the first BOQ being loaded, False for subsequent BOQs
            
        Returns:
            Dictionary with offer information or None if cancelled
        """
        try:
            from ui.offer_info_dialog import show_offer_info_dialog
            
            # Determine if this is the first BOQ or subsequent
            previous_info = self.previous_offer_info if not is_first_boq else None
            
            # Show the dialog
            offer_info = show_offer_info_dialog(self.root, is_first_boq, previous_info)
            
            if offer_info:
                # Store the offer information
                self.current_offer_info = offer_info
                self.current_offer_name = offer_info['supplier_name']  # Use supplier name as the offer name
                
                # Store as previous info for subsequent BOQs (only if this is the first BOQ)
                if is_first_boq:
                    self.previous_offer_info = offer_info.copy()
                
                return offer_info
            
            return None
            
        except ImportError:
            # Fallback to simple dialog if new dialog is not available
            import tkinter.simpledialog
            offer_name = tkinter.simpledialog.askstring("Offer Name", "Enter a name or label for this offer:", parent=self.root)
            if offer_name is not None and offer_name.strip() != "":
                simple_info = {
                    'supplier_name': offer_name.strip(),
                    'project_name': '',
                    'project_size': '',
                    'date': '2025-07-11',
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.current_offer_info = simple_info
                self.current_offer_name = offer_name.strip()
                return simple_info
            return None
    
    def _prompt_offer_name(self):
        """Legacy method for backward compatibility"""
        offer_info = self._prompt_offer_info(is_first_boq=True)
        return offer_info['supplier_name'] if offer_info else None

    def _save_analysis(self, tab):
        """Save the current analysis including DataFrame and mapping information as a pickle file."""
        df = getattr(tab, 'final_dataframe', None)
        if df is None:
            messagebox.showerror("Error", "No analysis to save.")
            return
        
        # print(f"[DEBUG] Saving analysis with DataFrame shape: {df.shape}")
        # print(f"[DEBUG] DataFrame columns: {df.columns.tolist()}")
        # print(f"[DEBUG] DataFrame first few rows:")
        print(df.head().to_string())
        
        if df.empty:
            messagebox.showerror("Error", "Analysis is empty - nothing to save.")
            return
            
        # Get the current offer name
        offer_name = self.current_offer_name if hasattr(self, 'current_offer_name') else None
        
        # Get mapping information if available
        mapping_data = None
        if hasattr(tab, 'stored_mapping_data') and tab.stored_mapping_data is not None:
            mapping_data = tab.stored_mapping_data
            # print(f"[DEBUG] Including stored mapping data with {len(mapping_data.get('sheets', []))} sheets")
        else:
            # Try to extract mapping data
            mapping_data = self._extract_mapping_from_tab(tab)
            if mapping_data:
                # print(f"[DEBUG] Extracted mapping data with {len(mapping_data.get('sheets', []))} sheets")
                pass
        
        # Get comparison information if available
        comparison_offers = getattr(tab, 'comparison_offers', None)
        is_comparison = self._is_comparison_dataset(df) if df is not None else False
        
        # Create a comprehensive save data dictionary
        save_data = {
            'dataframe': df,
            'offer_name': offer_name,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'mapping_data': mapping_data,
            'comparison_offers': comparison_offers,
            'is_comparison': is_comparison,
            'analysis_type': 'comparison' if is_comparison else 'single_offer'
        }
        
        file_path = filedialog.asksaveasfilename(
            title="Save Analysis as Pickle",
            defaultextension=".pkl",
            filetypes=[("Pickle files", "*.pkl"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'wb') as f:
                    pickle.dump(save_data, f)
                # print(f"[DEBUG] Successfully saved analysis to {file_path}")
                
                # Show detailed success message
                details = [f"DataFrame: {df.shape[0]} rows, {df.shape[1]} columns"]
                if mapping_data:
                    details.append(f"Mapping: {len(mapping_data.get('sheets', []))} sheets")
                if comparison_offers:
                    details.append(f"Comparison: {len(comparison_offers)} offers")
                
                success_msg = f"Analysis saved successfully!\n\n" + "\n".join(details)
                messagebox.showinfo("Success", success_msg)
                self._update_status(f"Analysis saved to: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save analysis: {str(e)}")
                # print(f"[DEBUG] Failed to save: {e}")
                import traceback
                traceback.print_exc()

    def _save_mappings(self, tab):
        """Save the current mappings (sheet, column, row) as a pickle file."""
        # Use the robust extraction method to get mappings from any state (new, loaded, compared)
        mappings_data = self._extract_mapping_from_tab(tab)

        if mappings_data is None:
            messagebox.showerror("Error", "No mappings found to save.")
            return

        # Ensure the final DataFrame in the mapping data is the most current version from the UI
        if hasattr(tab, 'final_dataframe') and tab.final_dataframe is not None:
            mappings_data['final_dataframe'] = tab.final_dataframe.copy()
            # print(f"[DEBUG] Saving updated final DataFrame with shape: {tab.final_dataframe.shape}")
        
        # Add the current offer name, which might have been updated
        if hasattr(self, 'current_offer_name') and self.current_offer_name:
            mappings_data['offer_name'] = self.current_offer_name
        
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
                
                # print(f"[DEBUG] Loaded data type: {type(loaded_data)}")
                
                # Handle both old format (just DataFrame) and new format (dictionary)
                if isinstance(loaded_data, dict):
                    df = loaded_data.get('dataframe')
                    self.current_offer_name = loaded_data.get('offer_name')
                    mapping_data = loaded_data.get('mapping_data')
                    comparison_offers = loaded_data.get('comparison_offers')
                    is_comparison = loaded_data.get('is_comparison', False)
                    analysis_type = loaded_data.get('analysis_type', 'unknown')
                    
                    # print(f"[DEBUG] Enhanced format - DataFrame shape: {df.shape if df is not None else 'None'}")
                    # print(f"[DEBUG] Offer name: {self.current_offer_name}")
                    # print(f"[DEBUG] Has mapping data: {mapping_data is not None}")
                    # print(f"[DEBUG] Comparison offers: {comparison_offers}")
                    # print(f"[DEBUG] Is comparison: {is_comparison}")
                else:
                    # Legacy format - just DataFrame
                    df = loaded_data
                    self.current_offer_name = None
                    mapping_data = None
                    comparison_offers = None
                    is_comparison = False
                    analysis_type = 'legacy'
                    # print(f"[DEBUG] Legacy format - DataFrame shape: {df.shape if df is not None else 'None'}")
                
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # print(f"[DEBUG] DataFrame columns: {df.columns.tolist()}")
                    # print(f"[DEBUG] DataFrame first few rows:")
                    print(df.head().to_string())
                    
                    # Create a new tab for the loaded analysis
                    filename = os.path.basename(file_path)
                    tab_name = f"Loaded: {filename}"
                    if analysis_type == 'comparison':
                        tab_name = f"Comparison: {filename}"
                    
                    tab = ttk.Frame(self.notebook)
                    self.notebook.add(tab, text=tab_name)
                    self.notebook.select(tab)
                    
                    # Store the loaded data in the tab
                    if mapping_data:
                        # Validate and fix position data integrity
                        self._validate_position_data_integrity(mapping_data)
                        tab.stored_mapping_data = mapping_data
                        # print(f"[DEBUG] Stored mapping data with {len(mapping_data.get('sheets', []))} sheets")
                    
                    if comparison_offers:
                        tab.comparison_offers = comparison_offers
                        # print(f"[DEBUG] Stored comparison offers: {comparison_offers}")
                    
                    # Show the data in the grid
                    self._show_final_categorized_data(tab, df, None)
                    
                    # Update status with detailed information
                    details = [f"DataFrame: {df.shape[0]} rows, {df.shape[1]} columns"]
                    if mapping_data:
                        details.append(f"Mapping: {len(mapping_data.get('sheets', []))} sheets")
                    if comparison_offers:
                        details.append(f"Comparison: {len(comparison_offers)} offers")
                    
                    status_msg = f"Analysis loaded - " + ", ".join(details)
                    self._update_status(status_msg)
                    
                elif isinstance(df, pd.DataFrame) and df.empty:
                    messagebox.showerror("Error", "The loaded analysis contains an empty DataFrame")
                    # print("[DEBUG] DataFrame is empty")
                else:
                    messagebox.showerror("Error", "Invalid analysis file format - no DataFrame found")
                    # print(f"[DEBUG] Invalid format - df type: {type(df)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load analysis: {str(e)}")
                logger.error(f"Failed to load analysis from {file_path}: {e}", exc_info=True)
                # print(f"[DEBUG] Exception loading analysis: {e}")
                import traceback
                traceback.print_exc()

    def _compare_full(self, tab):
        """Handle loading and comparing another BOQ with identical structure"""
        try:
            # Get the current master dataset
            master_df = getattr(tab, 'final_dataframe', None)
            if master_df is None or master_df.empty:
                messagebox.showerror("Error", "No master dataset found. Please load and categorize a BOQ first.")
                return
            
            # Get the current offer name (for the master dataset)
            master_offer_name = self.current_offer_name if hasattr(self, 'current_offer_name') and self.current_offer_name else "Offer1"
            
            # Check if this is the first comparison (need to restructure master dataset)
            if not self._is_comparison_dataset(master_df):
                # print("[DEBUG] Converting master dataset to comparison format")
                
                # Extract and store mapping data before converting to comparison format
                master_mapping = self._extract_mapping_from_tab(tab)
                if master_mapping is None:
                    messagebox.showerror("Error", "Could not extract mapping from master dataset. Please ensure the dataset was created with saved mappings.")
                    return
                
                master_df = self._convert_to_comparison_format(master_df, master_offer_name)
                tab.final_dataframe = master_df
                tab.master_offer_name = master_offer_name
                tab.comparison_offers = [master_offer_name]
                
                # Reset summary tree to ensure it gets rebuilt with new structure
                tab.summary_tree = None
                
                # Ensure the mapping data is stored in the tab for future comparisons
                if not hasattr(tab, 'stored_mapping_data'):
                    tab.stored_mapping_data = master_mapping
                
                # Update the display
                self._update_comparison_display(tab, master_df)
            
            # Prompt for new offer information (subsequent BOQ)
            new_offer_info = self._prompt_offer_info(is_first_boq=False)
            if new_offer_info is None:
                self._update_status("Comparison cancelled (no offer information provided).")
                return
            
            new_offer_name = new_offer_info['supplier_name']
            
            # Check for duplicate offer names
            existing_offers = getattr(tab, 'comparison_offers', [])
            if new_offer_name in existing_offers:
                messagebox.showerror("Error", f"Offer name '{new_offer_name}' already exists. Please choose a different name.")
                return
            
            # Get the saved mapping from the master dataset
            master_mapping = self._extract_mapping_from_tab(tab)
            if master_mapping is None:
                messagebox.showerror("Error", "Could not extract mapping from master dataset. Please ensure the dataset was created with saved mappings.")
                return
            
            self._update_status(f"Master mapping extracted. Please select BOQ file for '{new_offer_name}'.")
            
            # Select BOQ file to compare
            filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            filename = filedialog.askopenfilename(title=f"Select BOQ File for {new_offer_name}", filetypes=filetypes)
            if not filename:
                self._update_status("Comparison cancelled (no file selected).")
                return
            
            # Process the new file with the master mapping
            self._process_comparison_file(tab, filename, new_offer_name, master_mapping)
            
        except Exception as e:
            # print(f"[DEBUG] Error in _compare_full: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to start comparison: {str(e)}")

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
            
            # Validate and fix position data integrity
            self._validate_position_data_integrity(mapping_data)
                
            self.saved_mapping = mapping_data
            # print(f"[DEBUG] Loaded mapping with {len(mapping_data['sheets'])} sheets")
            
            # Debug: Check if mapping contains final categorized data
            if 'final_dataframe' in mapping_data:
                df = mapping_data['final_dataframe']
                # print(f"[DEBUG] Mapping contains final DataFrame with shape: {df.shape if df is not None else None}")
                if df is not None and 'category' in df.columns:
                    # print(f"[DEBUG] Final DataFrame has category column with {len(df)} rows")
                    pass
                else:
                    # print("[DEBUG] Final DataFrame missing category column")
                    pass
            else:
                # print("[DEBUG] Mapping does NOT contain final DataFrame - will require categorization")
            
            # Step 2: Prompt for BOQ file to analyze
                pass
            self._update_status("Mapping loaded. Please select a BOQ file to analyze.")
            
            # Prompt for offer information
            offer_info = self._prompt_offer_info(is_first_boq=True)
            if offer_info is None:
                self._update_status("Analysis cancelled (no offer information provided).")
                return
            
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
            # print(f"[DEBUG] Error loading mapping: {e}")
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
                
                # print(f"[DEBUG] File has sheets: {visible_sheets}")
                # print(f"[DEBUG] Mapping expects sheets: {[s['sheet_name'] for s in mapping_data['sheets']]}")
                
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
                # print(f"[DEBUG] Error in process_with_mapping: {e}")
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: self._on_processing_error(tab, filename, loading_label))
        
        threading.Thread(target=process_with_mapping, daemon=True).start()

    def _apply_saved_mappings(self, tab, file_mapping, mapping_data, loading_widget):
        """Apply saved column and row mappings to the processed file with strict validation"""
        try:
            # print("[DEBUG] Applying saved mappings with strict validation...")
            
            # Apply column mappings and validate structure
            for sheet in file_mapping.sheets:
                sheet_name = sheet.sheet_name
                
                # Find corresponding mapping
                saved_sheet = next((s for s in mapping_data['sheets'] if s['sheet_name'] == sheet_name), None)
                if not saved_sheet:
                    continue
                
                # print(f"[DEBUG] Validating structure for sheet: {sheet_name}")
                
                # Auto-align header text when column count and mapped types match but header strings differ.
                try:
                    saved_cm_list = saved_sheet.get('column_mappings', [])
                    if saved_cm_list and len(saved_cm_list) == len(getattr(sheet, 'column_mappings', [])):
                        headers_differ = False
                        mapped_types_match = True
                        for cur_cm, saved_cm in zip(sheet.column_mappings, saved_cm_list):
                            saved_mtype = saved_cm.get('mapped_type') if isinstance(saved_cm, dict) else getattr(saved_cm, 'mapped_type', None)
                            if saved_mtype != getattr(cur_cm, 'mapped_type', None):
                                mapped_types_match = False
                                break
                            saved_header = saved_cm.get('original_header') if isinstance(saved_cm, dict) else getattr(saved_cm, 'original_header', '')
                            if saved_header and saved_header.strip() != getattr(cur_cm, 'original_header', '').strip():
                                headers_differ = True

                        if mapped_types_match and headers_differ:
                            for cur_cm, saved_cm in zip(sheet.column_mappings, saved_cm_list):
                                saved_header = saved_cm.get('original_header') if isinstance(saved_cm, dict) else getattr(saved_cm, 'original_header', '')
                                cur_cm.original_header = saved_header
                            # print(f"[DEBUG] Auto-aligned headers for sheet: {sheet_name}")
                except Exception:
                    pass
                

                
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
                
                # print(f"[DEBUG] Sheet '{sheet_name}': Structure validated and mappings applied")
            
            # ENHANCED VALIDATION: Position-description validation for the entire file
            logger.info("Performing position-description validation for Use Mapping workflow...")
            position_validation = self._validate_position_description_match(file_mapping, mapping_data)
            
            if not position_validation['is_valid']:
                loading_widget.destroy()
                try:
                    _handle_validation_failure(position_validation, "Use Mapping", "position-description validation")
                except PositionValidationError:
                    # Process terminated due to validation failure
                    return
            
            logger.info(f"Use Mapping position-description validation passed: {position_validation['summary']}")
            
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
            # print(f"[DEBUG] Error applying mappings: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to apply mappings: {str(e)}")
            loading_widget.destroy()

    def _validate_exact_column_structure(self, sheet, saved_sheet):
        """Validate that the column structure matches exactly (original headers and order)."""
        try:
            # Current original headers extracted from ColumnMappings
            current_headers = [getattr(cm, 'original_header', '') for cm in getattr(sheet, 'column_mappings', [])]

            # Saved original headers extracted from saved mapping data
            saved_headers = []
            for saved_mapping in saved_sheet.get('column_mappings', []):
                if isinstance(saved_mapping, dict):
                    saved_headers.append(saved_mapping.get('original_header', ''))

            # Must have identical length and identical header strings in order
            if len(current_headers) != len(saved_headers):
                return False

            for cur, saved in zip(current_headers, saved_headers):
                if cur.strip() != saved.strip():
                    return False

            return True
        except Exception:
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
            
            # print(f"[DEBUG] Saved mapping has {len(saved_valid_rows)} valid rows at indices: {saved_valid_indices[:10]}...")
            
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
            
            # print(f"[DEBUG] Current file has {len(current_row_data)} rows at the expected valid indices")
            
            # Must have same number of valid rows
            if len(current_row_data) != len(saved_valid_rows):
                # print(f"[DEBUG] Valid row count mismatch: {len(current_row_data)} vs {len(saved_valid_rows)}")
                # print(f"[DEBUG] Missing indices: {set(saved_valid_indices) - {idx for idx, _ in current_valid_rows}}")
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
                        # print(f"[DEBUG] Valid row {row_idx} mismatch:")
                        print(f"  Current: {current_normalized[:3]}...")
                        print(f"  Saved:   {saved_normalized[:3]}...")
            
            if mismatched_rows > 0:
                # print(f"[DEBUG] {mismatched_rows} valid rows don't match exactly")
                return False
            
            # print("[DEBUG] Row structure validation passed - all valid rows match")
            return True
            
        except Exception as e:
            # print(f"[DEBUG] Error validating row structure: {e}")
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
                    # print(f"[DEBUG] Applied exact mapping: {original_header} -> {cm.mapped_type}")
            
        except Exception as e:
            # print(f"[DEBUG] Error applying column mappings: {e}")

            pass
    def _apply_exact_row_classifications(self, sheet, saved_sheet):
        """Apply the saved row classifications exactly"""
        try:
            saved_classifications = saved_sheet.get('row_classifications', [])
            
            # Initialize row validity for this sheet
            sheet_name = sheet.sheet_name
            if not hasattr(self, 'row_validity'):
                self.row_validity = {}
            self.row_validity[sheet_name] = {}
            
            # Create a mapping of row_index to validity from saved data
            saved_validity_map = {}
            for saved_rc in saved_classifications:
                if isinstance(saved_rc, dict):
                    saved_row_type = saved_rc.get('row_type', '')
                    is_valid = saved_row_type in ['primary_line_item', 'PRIMARY_LINE_ITEM']
                    row_index = saved_rc.get('row_index', -1)
                    if row_index >= 0:
                        saved_validity_map[row_index] = is_valid
            
            # Apply saved validity to corresponding rows by matching row_index
            if hasattr(sheet, 'row_classifications'):
                for current_rc in sheet.row_classifications:
                    row_index = getattr(current_rc, 'row_index', -1)
                    if row_index >= 0 and row_index in saved_validity_map:
                        self.row_validity[sheet_name][row_index] = saved_validity_map[row_index]
                    else:
                        # Default to False for rows not found in saved mapping
                        self.row_validity[sheet_name][row_index] = False
            
            print(f"[DEBUG] Applied row validity for {len(self.row_validity[sheet_name])} rows in sheet '{sheet_name}'")
            valid_count = sum(1 for v in self.row_validity[sheet_name].values() if v)
            print(f"[DEBUG] {valid_count} rows are valid, {len(self.row_validity[sheet_name]) - valid_count} rows are invalid")
             
        except Exception as e:
            # Log error but don't fail the entire process
            pass

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
            # print(f"[DEBUG] Error showing mapped file review: {e}")
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
            base_required_types = [col_type.value for col_type in self.column_mapper.config.get_required_columns()]
        else:
            base_required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
        
        # Add new columns as "required" if they are successfully mapped
        required_types = base_required_types.copy()
        if hasattr(sheet, 'column_mappings'):
            for mapping in sheet.column_mappings:
                mapped_type = getattr(mapping, 'mapped_type', 'unknown')
                confidence = getattr(mapping, 'confidence', 0)
                if mapped_type in ['scope', 'manhours', 'wage'] and confidence > 0:
                    if mapped_type not in required_types:
                        required_types.append(mapped_type)
        
        # Display column order (include new columns)
        display_column_order = ["code", "description", "unit", "quantity", "unit_price", "total_price", "scope", "manhours", "wage"]
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
                tree.column(col, width=120 if col != "#" else 40, anchor=tk.W, minwidth=50, stretch=False)
        
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
                    
                    # Format numbers - ensure consistent formatting with other row review
                    if col in ['unit_price', 'total_price', 'wage']:
                        val = format_number(val, is_currency=True)
                    elif col in ['quantity', 'manhours']:
                        val = format_number(val, is_currency=False)
                        # Special formatting for manhours - only 2 decimals
                        if col == 'manhours' and val and val != '':
                            try:
                                num_val = float(str(val).replace(',', '.'))
                                val = f"{num_val:.2f}"
                            except:
                                pass
                    
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

    # ===== COMPARISON FUNCTIONALITY =====
    
    def _is_comparison_dataset(self, df):
        """Check if the DataFrame is already in comparison format (has offer-specific columns)"""
        if df is None or df.empty:
            return False
        
        # Check if any column names contain offer-specific suffixes (both old and new formats)
        for col in df.columns:
            if any(col.startswith(prefix) for prefix in ['quantity_', 'unit_price_', 'total_price_', 'manhours_', 'wage_', 'quantity[', 'unit_price[', 'total_price[', 'manhours[', 'wage[']):
                return True
        return False
    
    def _convert_to_comparison_format(self, df, offer_name):
        """Convert a single-offer DataFrame to comparison format"""
        import pandas as pd
        
        # Create a copy to avoid modifying the original
        comparison_df = df.copy()
        
        # Rename the price/quantity columns to include offer name in square brackets
        column_mapping = {
            'quantity': f'quantity[{offer_name}]',
            'unit_price': f'unit_price[{offer_name}]',
            'total_price': f'total_price[{offer_name}]',
            'manhours': f'manhours[{offer_name}]',
            'wage': f'wage[{offer_name}]'
        }
        
        comparison_df = comparison_df.rename(columns=column_mapping)
        
        # Reorder columns according to the specified order: sheet, category, code, description, unit, quantity[Offer1], unit_price[Offer1], total_price[Offer1], scope, manhours[Offer1], wage[Offer1]
        base_columns = ['sheet', 'category', 'code', 'description', 'unit']
        offer_columns = [f'quantity[{offer_name}]', f'unit_price[{offer_name}]', f'total_price[{offer_name}]']
        
        # Add new columns if they exist (scope stays as-is, manhours and wage get offer suffix)
        if 'scope' in comparison_df.columns:
            base_columns.append('scope')
        if f'manhours[{offer_name}]' in comparison_df.columns:
            offer_columns.append(f'manhours[{offer_name}]')
        if f'wage[{offer_name}]' in comparison_df.columns:
            offer_columns.append(f'wage[{offer_name}]')
        
        # Only include columns that exist in the DataFrame
        ordered_columns = []
        for col in base_columns + offer_columns:
            if col in comparison_df.columns:
                ordered_columns.append(col)
        
        # Add any remaining columns that weren't in the predefined order
        for col in comparison_df.columns:
            if col not in ordered_columns:
                ordered_columns.append(col)
        
        comparison_df = comparison_df[ordered_columns]
        
        # print(f"[DEBUG] Converted to comparison format with columns: {comparison_df.columns.tolist()}")
        return comparison_df
    
    def _update_comparison_display(self, tab, comparison_df):
        """Update the display to show the comparison DataFrame"""
        # Debug: Check if unit column exists in comparison DataFrame
        # print(f"[DEBUG] _update_comparison_display: DataFrame columns: {comparison_df.columns.tolist()}")
        if 'unit' in comparison_df.columns:
            unit_values = comparison_df['unit'].value_counts()
            empty_units = comparison_df['unit'].isna().sum() + (comparison_df['unit'] == '').sum()
            total_rows = len(comparison_df)
            # print(f"[DEBUG] Unit column: {total_rows - empty_units}/{total_rows} values populated")
            if total_rows - empty_units > 0:
                # print(f"[DEBUG] ✅ Unit values found: {unit_values.to_dict()}")
                pass
            else:
                # print(f"[DEBUG] ❌ All unit values are empty!")
                pass
        else:
            # print(f"[DEBUG] ❌ Unit column missing from comparison DataFrame!")
        
        # Update the treeview with new columns
            pass
        if hasattr(tab, 'final_data_tree'):
            tree = tab.final_data_tree
            
            # Store current column widths to preserve user changes
            current_widths = {}
            if hasattr(tree, 'column'):
                try:
                    current_columns = tree['columns'] if tree['columns'] else []
                    for col in current_columns:
                        try:
                            current_widths[col] = tree.column(col, 'width')
                        except:
                            pass
                except:
                    pass
            
            # Also check if tab has stored column widths from previous user interactions
            if hasattr(tab, 'user_column_widths'):
                for col, width in tab.user_column_widths.items():
                    current_widths[col] = width
                # print(f"[DEBUG] Restored {len(tab.user_column_widths)} user column widths from tab")
            
            # Initialize user column widths storage if not exists
            if not hasattr(tab, 'user_column_widths'):
                tab.user_column_widths = {}
            
            # CRITICAL FIX: Ensure unit column is present like in single BOQ mode
            new_columns = list(comparison_df.columns)
            
            # Enforce required columns similar to _build_final_grid_dataframe
            required_base_columns = ['sheet', 'category', 'code', 'description', 'unit']
            for req_col in required_base_columns:
                if req_col not in new_columns and req_col not in comparison_df.columns:
                    # print(f"[DEBUG] Adding missing required column: {req_col}")
                    comparison_df[req_col] = ''  # Add empty column
                    new_columns.append(req_col)
                elif req_col in comparison_df.columns and req_col not in new_columns:
                    # print(f"[DEBUG] Adding existing column to display: {req_col}")
                    new_columns.append(req_col)
            
            # CRITICAL FIX: Use DataFrame column order directly to avoid column misalignment
            # The TreeView display columns MUST match the DataFrame column order exactly
            # to prevent values from being displayed in wrong columns
            new_columns = comparison_df.columns.tolist()
            tree['columns'] = new_columns
            
            # Update column headings and widths
            for col in new_columns:
                # Format column names for display (handle both old and new formats)
                if '[' in col and ']' in col:
                    # New format: prefix[offer] -> Prefix [Offer]
                    prefix, suffix = col.split('[', 1)
                    offer_name = suffix.rstrip(']')
                    display_name = f"{prefix.replace('_', ' ').title()} [{offer_name}]"
                else:
                    # Old format: prefix_offer -> Prefix Offer
                    display_name = col.replace('_', ' ').title()
                
                tree.heading(col, text=display_name)
                
                # Determine column width - preserve user changes or auto-size
                if col in current_widths:
                    # Use previously set width (user may have resized)
                    width = current_widths[col]
                    # print(f"[DEBUG] Preserving user width for {col}: {width}px")
                elif col == 'description':
                    # Description column gets fixed width due to long text
                    width = 250
                else:
                    # Auto-size based on content and header
                    width = self._calculate_column_width(comparison_df, col, display_name)
                
                # CRITICAL: Set column properties with stretch=False to prevent auto-fitting
                # Use smaller minwidth to allow better narrowing
                if col.startswith(('quantity_', 'unit_price_', 'total_price_', 'manhours_', 'wage_', 'quantity[', 'unit_price[', 'total_price[', 'manhours[', 'wage[')):
                    tree.column(col, width=width, minwidth=50, anchor=tk.E, stretch=False)
                else:
                    tree.column(col, width=width, minwidth=50, anchor=tk.W, stretch=False)
            
            # Store current widths in tab for future reference
            for col in new_columns:
                try:
                    tab.user_column_widths[col] = tree.column(col, 'width')
                except:
                    pass
            
            # Repopulate the treeview
            self._populate_final_data_treeview(tree, comparison_df, new_columns)
            
            # CRITICAL FIX: Update controller.current_files with the converted DataFrame
            # Find the file entry in controller.current_files and update it
            for file_key, file_data in self.controller.current_files.items():
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == tab:
                    # Update the categorized_dataframe in the file mapping
                    file_data['file_mapping'].categorized_dataframe = comparison_df
                    # Also update the file_data directly
                    file_data['categorized_dataframe'] = comparison_df
                    print(f"[DEBUG] Updated controller.current_files[{file_key}] with converted DataFrame")
                    print(f"[DEBUG] Converted DataFrame columns: {list(comparison_df.columns)}")
                    break
            
            # Reset summary tree so it gets rebuilt with the new DataFrame structure
            tab.summary_tree = None
            
            # Refresh the summary grid after comparison
            print(f"[DEBUG] Calling centralized summary grid refresh after comparison")
            self._refresh_summary_grid_centralized()
            
            # Remove any existing resize bindings to avoid conflicts
            try:
                tree.unbind('<Button-1>')
                tree.unbind('<ButtonRelease-1>')
                tree.unbind('<B1-Motion>')
                tree.unbind('<Configure>')
                tree.unbind('<Motion>')
            except:
                pass  # Events might not be bound yet
            
            # Prevent the treeview from auto-adjusting column widths
            def prevent_auto_resize(event=None):
                """Prevent automatic column width adjustments"""
                # Don't interfere during active resizing
                if resize_state['prevent_auto_resize'] or resize_state['resizing']:
                    return "break"
                    
                if hasattr(tab, 'user_column_widths'):
                    for col, width in tab.user_column_widths.items():
                        try:
                            current_width = tree.column(col, 'width')
                            if current_width != width and abs(current_width - width) > 5:  # Only restore if significant difference
                                tree.column(col, width=width, minwidth=50)
                                # print(f"[DEBUG] Restored column {col} width to {width}px (was {current_width}px)")
                        except:
                            pass
                return "break"  # Prevent further event processing
            
            # Track resize state - make it persistent across BOQ loads
            if not hasattr(tab, 'resize_state'):
                tab.resize_state = {
                    'resizing': False, 
                    'resize_column': None, 
                    'start_x': 0, 
                    'original_width': 0,
                    'prevent_auto_resize': False
                }
            resize_state = tab.resize_state
            
            def on_button_press(event):
                """Handle mouse button press - check if we're starting a resize"""
                try:
                    region = tree.identify_region(event.x, event.y)
                    if region == "separator":
                        # Find which column we're resizing
                        col = tree.identify_column(event.x)
                        if col and col.startswith('#'):
                            col_index = int(col[1:]) - 1  # Convert #1, #2, etc to 0, 1, etc
                            if 0 <= col_index < len(new_columns):
                                resize_state['resizing'] = True
                                resize_state['resize_column'] = new_columns[col_index]
                                resize_state['start_x'] = event.x
                                resize_state['original_width'] = tree.column(resize_state['resize_column'], 'width')
                                resize_state['prevent_auto_resize'] = True
                                # print(f"[DEBUG] Started resizing column {resize_state['resize_column']} from {resize_state['original_width']}px")
                                return "break"  # Prevent default behavior
                except Exception as e:
                    # print(f"[DEBUG] Error in button press: {e}")
                
                # Not a resize operation, allow normal processing
                    pass
                resize_state['prevent_auto_resize'] = False
                return None
            
            def on_button_motion(event):
                """Handle mouse motion during resize"""
                if resize_state['resizing'] and resize_state['resize_column']:
                    try:
                        # Calculate new width based on mouse movement
                        delta = event.x - resize_state['start_x']
                        new_width = max(50, resize_state['original_width'] + delta)  # Lower minimum for narrowing
                        
                        # Temporarily disable auto-resize prevention during manual resize
                        resize_state['prevent_auto_resize'] = True
                        
                        # Apply the new width immediately with explicit override
                        tree.column(resize_state['resize_column'], width=new_width, minwidth=50)
                        
                        # Update stored width immediately to prevent snap-back
                        if not hasattr(tab, 'user_column_widths'):
                            tab.user_column_widths = {}
                        tab.user_column_widths[resize_state['resize_column']] = new_width
                        
                        # print(f"[DEBUG] Resizing column {resize_state['resize_column']} to {new_width}px")
                        return "break"
                    except Exception as e:
                        # print(f"[DEBUG] Error in motion: {e}")
                        pass
                return None 
            
            def on_button_release(event):
                """Handle mouse button release - finalize resize"""
                if resize_state['resizing'] and resize_state['resize_column']:
                    try:
                        col = resize_state['resize_column']
                        final_width = tree.column(col, 'width')
                        
                        # Store the user's preferred width
                        if not hasattr(tab, 'user_column_widths'):
                            tab.user_column_widths = {}
                        tab.user_column_widths[col] = final_width
                        
                        # print(f"[DEBUG] Finished resizing column {col} to {final_width}px - saved preference")
                        
                        # Reset resize state
                        resize_state['resizing'] = False
                        resize_state['resize_column'] = None
                        resize_state['start_x'] = 0
                        resize_state['original_width'] = 0
                        resize_state['prevent_auto_resize'] = False
                        
                        return "break"
                    except Exception as e:
                        # print(f"[DEBUG] Error in button release: {e}")
                
                        pass
                resize_state['prevent_auto_resize'] = False
                return None
            
            # Bind the events with proper return values to control event propagation
            tree.bind('<Button-1>', on_button_press)
            tree.bind('<B1-Motion>', on_button_motion)
            tree.bind('<ButtonRelease-1>', on_button_release)
            
            # Prevent configure events from auto-resizing columns
            tree.bind('<Configure>', prevent_auto_resize)
            
            # Override the default column adjustment behavior
            def override_column_width(col, option=None, **kw):
                """Override column width setting to prevent auto-adjustments"""
                if option == 'width' and resize_state['prevent_auto_resize']:
                    # During resize, allow the width change
                    return tree.tk.call(tree._w, 'column', col, '-width', kw.get('width', kw.get('w', 100)))
                elif option == 'width' and col in tab.user_column_widths:
                    # Use stored user width
                    return tree.tk.call(tree._w, 'column', col, '-width', tab.user_column_widths[col])
                else:
                    # Default behavior for other options
                    return tree.tk.call(tree._w, 'column', col, f'-{option}' if option else '', *kw.values())
            
            # Apply the override (this is a bit hacky but necessary for Tkinter)
            # Only patch if not already patched
            if not hasattr(tree, '_column_patched'):
                original_column = tree.column
                def patched_column(col, option=None, **kw):
                    if option == 'width' and not kw and not resize_state['prevent_auto_resize'] and hasattr(tab, 'user_column_widths') and col in tab.user_column_widths:
                        # Return stored user-defined width (GET operation)
                        return tab.user_column_widths[col]
                    else:
                        return original_column(col, option, **kw)
                
                tree.column = patched_column
                tree._column_patched = True
        
        # print(f"[DEBUG] Updated comparison display with {len(comparison_df)} rows and {len(comparison_df.columns)} columns")
    
    def _calculate_column_width(self, df, column_name, display_name):
        """Calculate optimal column width based on content and header"""
        import tkinter.font as tkFont
        
        try:
            # Get default font for measurements
            font = tkFont.nametofont("TkDefaultFont")
        except:
            # Fallback if font not available
            # Approximate character width
            header_width = len(display_name) * 8
            max_content_width = 0
            
            if column_name in df.columns:
                sample_size = min(100, len(df))
                if sample_size > 0:
                    sample_data = df[column_name].head(sample_size)
                    for value in sample_data:
                        if value is not None and str(value).strip():
                            content_width = len(str(value)) * 8
                            max_content_width = max(max_content_width, content_width)
            
            optimal_width = max(header_width, max_content_width) + 30  # Extra padding
            return max(100, min(optimal_width, 400))
        
        # Measure header width with proper font
        header_width = font.measure(display_name)
        # print(f"[DEBUG] Header '{display_name}' width: {header_width}px")
        
        # Check content width (sample first 100 rows for performance)
        max_content_width = 0
        sample_size = min(100, len(df))
        
        if column_name in df.columns and sample_size > 0:
            sample_data = df[column_name].head(sample_size)
            for value in sample_data:
                if value is not None and str(value).strip():
                    content_width = font.measure(str(value))
                    max_content_width = max(max_content_width, content_width)
        
        # print(f"[DEBUG] Column '{column_name}' max content width: {max_content_width}px")
        
        # Calculate optimal width (ensure header is fully visible + padding)
        optimal_width = max(header_width + 30, max_content_width + 20)  # Extra padding for header
        
        # Apply reasonable limits
        min_width = 100  # Increased minimum to ensure headers show
        max_width = 400  # Increased maximum for better visibility
        
        final_width = max(min_width, min(optimal_width, max_width))
        # print(f"[DEBUG] Column '{column_name}' final width: {final_width}px")
        
        return final_width
    
    def _extract_mapping_from_tab(self, tab):
        """Extract mapping information from the current tab for reuse"""
        # First, check if we already have stored mapping data in the tab (for comparison datasets)
        if hasattr(tab, 'stored_mapping_data') and tab.stored_mapping_data is not None:
            # print("[DEBUG] Using stored mapping data from tab")
            return tab.stored_mapping_data
        
        # Try to find the file mapping for this tab (for original datasets)
        for file_data in self.controller.current_files.values():
            if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == tab:
                # Extract the mapping information we need
                file_mapping = file_data['file_mapping']
                
                mapping_data = {
                    'sheets': [],
                    'row_validity': getattr(self, 'row_validity', {}),
                }
                
                # Save relevant info from each sheet
                for sheet in getattr(file_mapping, 'sheets', []):
                    sheet_info = {
                        'sheet_name': getattr(sheet, 'sheet_name', None),
                        'sheet_type': getattr(sheet, 'sheet_type', None),
                        'header_row_index': getattr(sheet, 'header_row_index', None),
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
                    mapping_data['sheets'].append(sheet_info)
                
                # Add the final dataframe for category mapping
                if hasattr(tab, 'final_dataframe') and tab.final_dataframe is not None:
                    mapping_data['final_dataframe'] = tab.final_dataframe.copy()
                
                # Store the mapping data in the tab for future use
                tab.stored_mapping_data = mapping_data
                # print("[DEBUG] Extracted and stored mapping data from file_mapping")
                return mapping_data
        
        # If no file mapping found, return None
        # print("[DEBUG] Could not extract mapping from tab - no file mapping found")
        return None
    
    def _process_comparison_file(self, master_tab, filepath, new_offer_name, master_mapping):
        """Process a new BOQ file for comparison using the master mapping"""
        import os
        from pathlib import Path
        self._update_status(f"Processing {os.path.basename(filepath)} for comparison...")
        
        def process_in_thread():
            try:
                # Step 1: Load the file and get sheets
                from core.file_processor import ExcelProcessor
                processor = ExcelProcessor()
                processor.load_file(Path(filepath))
                visible_sheets = processor.get_visible_sheets()
                
                # print(f"[DEBUG] Comparison file has sheets: {visible_sheets}")
                # print(f"[DEBUG] Master mapping expects sheets: {[s['sheet_name'] for s in master_mapping['sheets']]}")
                
                # Step 2: Validate sheets match exactly
                mapping_sheets = {s['sheet_name'] for s in master_mapping['sheets']}
                file_sheets = set(visible_sheets)
                
                missing_sheets = mapping_sheets - file_sheets
                if missing_sheets:
                    error_msg = f"Missing sheets in comparison file: {', '.join(missing_sheets)}"
                    self.root.after(0, lambda: messagebox.showerror("Sheet Validation Error", error_msg))
                    return
                
                # Step 3: Process the file (basic processing first)
                file_mapping = self.controller.process_file(
                    Path(filepath),
                    progress_callback=lambda p, m: self.root.after(0, self.update_progress, p, m),
                    sheet_filter=list(mapping_sheets),
                    sheet_types={sheet: "BOQ" for sheet in mapping_sheets}
                )
                
                # Step 4: Apply the master mapping and merge with comparison dataset
                self.root.after(0, lambda: self._apply_mapping_and_merge(master_tab, file_mapping, master_mapping, new_offer_name))
                
            except Exception as e:
                # print(f"[DEBUG] Error in process_comparison_file: {e}")
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to process comparison file: {str(e)}"))
        
        threading.Thread(target=process_in_thread, daemon=True).start()
    
    def _apply_mapping_and_merge(self, master_tab, file_mapping, master_mapping, new_offer_name):
        """Apply master mapping to new file and merge with comparison dataset"""
        try:
            # print(f"[DEBUG] Applying master mapping and merging for offer: {new_offer_name}")
            
            # ENHANCED VALIDATION: Position-description validation before applying mappings
            logger.info("Performing position-description validation for Compare Full workflow...")
            position_validation = self._validate_position_description_match(file_mapping, master_mapping)
            
            if not position_validation['is_valid']:
                try:
                    _handle_validation_failure(position_validation, "Compare Full", "position-description validation")
                except PositionValidationError:
                    # Process terminated due to validation failure
                    return
            
            logger.info(f"Compare Full position-description validation passed: {position_validation['summary']}")
            
            # Apply the master mapping to the new file (reuse existing logic)
            # This is similar to _apply_saved_mappings but without the UI parts
            for sheet in file_mapping.sheets:
                sheet_name = sheet.sheet_name
                
                # Find corresponding mapping
                saved_sheet = next((s for s in master_mapping['sheets'] if s['sheet_name'] == sheet_name), None)
                if not saved_sheet:
                    continue
                
                # Apply the mappings
                self._apply_exact_column_mappings(sheet, saved_sheet)
                self._apply_exact_row_classifications(sheet, saved_sheet)
            
            # Build DataFrame from the new file
            new_df = self._build_dataframe_from_mapping(file_mapping, master_mapping)
            
            # Debug: Check if new DataFrame has unit column with values
            if 'unit' in new_df.columns:
                unit_values = new_df['unit'].value_counts()
                # print(f"[DEBUG] New DataFrame unit values: {unit_values.to_dict()}")
                empty_units = (new_df['unit'] == '') | (new_df['unit'].isna())
                if empty_units.any():
                    # print(f"[DEBUG] WARNING: New DataFrame has {empty_units.sum()} empty unit values!")
                    pass
            else:
                # print(f"[DEBUG] CRITICAL: New DataFrame missing unit column!")
            
            # Apply categories from master mapping
                pass
            new_df = self._apply_categories_from_mapping(new_df, master_mapping)
            
            # Merge with the master comparison dataset
            master_df = master_tab.final_dataframe
            
            # Debug: Check master DF state before merge
            # print(f"[DEBUG] Before merge - Master DF shape: {master_df.shape}")
            # print(f"[DEBUG] Before merge - Master DF columns: {master_df.columns.tolist()}")
            if 'unit' in master_df.columns:
                master_unit_values = master_df['unit'].value_counts()
                # print(f"[DEBUG] Before merge - Master DF unit values: {master_unit_values.to_dict()}")
            if 'total_price' in master_df.columns:
                # print(f"[DEBUG] Before merge - Master DF total_price sum: {master_df['total_price'].sum()}")
            
                pass
            merged_df = self._merge_comparison_datasets(master_df, new_df, new_offer_name)
            
            # Debug: Check master DF state after merge (should be unchanged)
            # print(f"[DEBUG] After merge - Master DF shape: {master_df.shape}")
            # print(f"[DEBUG] After merge - Master DF columns: {master_df.columns.tolist()}")
            if 'total_price' in master_df.columns:
                # print(f"[DEBUG] After merge - Master DF total_price sum: {master_df['total_price'].sum()}")
            # print(f"[DEBUG] After merge - Merged DF shape: {merged_df.shape}")
            # print(f"[DEBUG] After merge - Merged DF columns: {merged_df.columns.tolist()}")
            
            # Update the master tab
                pass
            master_tab.final_dataframe = merged_df
            if not hasattr(master_tab, 'comparison_offers'):
                master_tab.comparison_offers = []
            master_tab.comparison_offers.append(new_offer_name)
            
            # CRITICAL FIX: Update controller.current_files with the merged DataFrame
            # Find the file entry in controller.current_files and update it
            file_found = False
            for file_key, file_data in self.controller.current_files.items():
                # Try to match by tab reference first
                if hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab == master_tab:
                    file_found = True
                    print(f"[DEBUG] Found file by tab reference: {file_key}")
                # If tab reference is None, try to match by checking if this file has offers
                elif hasattr(file_data['file_mapping'], 'tab') and file_data['file_mapping'].tab is None:
                    if 'offers' in file_data and len(file_data['offers']) > 0:
                        file_found = True
                        print(f"[DEBUG] Found file by offers check: {list(file_data['offers'].keys())}")
                
                if file_found:
                    # Update the categorized_dataframe in the file mapping
                    file_data['file_mapping'].categorized_dataframe = merged_df
                    # Also update the file_data directly
                    file_data['categorized_dataframe'] = merged_df
                    
                    # CRITICAL FIX: Ensure the tab reference is maintained
                    file_data['file_mapping'].tab = master_tab
                    
                    # CRITICAL FIX: Store offer info for the new offer in the comparison
                    if 'offers' not in file_data:
                        file_data['offers'] = {}
                    
                    # Store the new offer info
                    new_offer_info = {
                        'supplier_name': new_offer_name,
                        'project_name': getattr(self, 'current_offer_info', {}).get('project_name', 'Unknown'),
                        'date': getattr(self, 'current_offer_info', {}).get('date', 'Unknown')
                    }
                    file_data['offers'][new_offer_name] = new_offer_info
                    
                    # CRITICAL FIX: Also store the first offer info if it's not already there
                    # The first offer should be the master offer (the one that was already loaded)
                    master_offer_name = getattr(master_tab, 'master_offer_name', None)
                    if master_offer_name and master_offer_name not in file_data['offers']:
                        master_offer_info = {
                            'supplier_name': master_offer_name,
                            'project_name': getattr(self, 'previous_offer_info', {}).get('project_name', 'Unknown'),
                            'date': getattr(self, 'previous_offer_info', {}).get('date', 'Unknown')
                        }
                        file_data['offers'][master_offer_name] = master_offer_info
                        print(f"[DEBUG] Added master offer info for {master_offer_name}: {master_offer_info}")
                    
                    print(f"[DEBUG] All offers after merge: {list(file_data['offers'].keys())}")
                    print(f"[DEBUG] Updated controller.current_files[{file_key}] with merged DataFrame")
                    print(f"[DEBUG] Merged DataFrame columns: {list(merged_df.columns)}")
                    print(f"[DEBUG] Added offer info for {new_offer_name}: {new_offer_info}")
                    print(f"[DEBUG] Maintained tab reference: {file_data['file_mapping'].tab}")
                    break
            
            # CRITICAL FIX: If no file found, use the first available file
            if not file_found:
                print(f"[DEBUG] WARNING: Could not find file_data to update with merged DataFrame!")
                print(f"[DEBUG] Available files: {list(self.controller.current_files.keys())}")
                for fd in self.controller.current_files.values():
                    if 'offers' in fd:
                        print(f"[DEBUG] File has offers: {list(fd['offers'].keys())}")
                    if 'file_mapping' in fd:
                        tab_ref = getattr(fd['file_mapping'], 'tab', None)
                        print(f"[DEBUG] File mapping tab: {tab_ref}")
                
                # Use the first available file as fallback
                if self.controller.current_files:
                    file_key = list(self.controller.current_files.keys())[0]
                    file_data = self.controller.current_files[file_key]
                    print(f"[DEBUG] Using fallback file: {file_key}")
                    
                    # Update the categorized_dataframe in the file mapping
                    file_data['file_mapping'].categorized_dataframe = merged_df
                    # Also update the file_data directly
                    file_data['categorized_dataframe'] = merged_df
                    
                    # CRITICAL FIX: Ensure the tab reference is maintained
                    file_data['file_mapping'].tab = master_tab
                    
                    # CRITICAL FIX: Store offer info for the new offer in the comparison
                    if 'offers' not in file_data:
                        file_data['offers'] = {}
                    
                    # Store the new offer info
                    new_offer_info = {
                        'supplier_name': new_offer_name,
                        'project_name': getattr(self, 'current_offer_info', {}).get('project_name', 'Unknown'),
                        'date': getattr(self, 'current_offer_info', {}).get('date', 'Unknown')
                    }
                    file_data['offers'][new_offer_name] = new_offer_info
                    
                    # CRITICAL FIX: Also store the first offer info if it's not already there
                    # The first offer should be the master offer (the one that was already loaded)
                    master_offer_name = getattr(master_tab, 'master_offer_name', None)
                    if master_offer_name and master_offer_name not in file_data['offers']:
                        master_offer_info = {
                            'supplier_name': master_offer_name,
                            'project_name': getattr(self, 'previous_offer_info', {}).get('project_name', 'Unknown'),
                            'date': getattr(self, 'previous_offer_info', {}).get('date', 'Unknown')
                        }
                        file_data['offers'][master_offer_name] = master_offer_info
                        print(f"[DEBUG] Added master offer info for {master_offer_name}: {master_offer_info}")
                    
                    print(f"[DEBUG] All offers after fallback merge: {list(file_data['offers'].keys())}")
                    print(f"[DEBUG] Updated fallback file with merged DataFrame")
                    print(f"[DEBUG] Added offer info for {new_offer_name}: {new_offer_info}")
                    print(f"[DEBUG] Maintained tab reference: {file_data['file_mapping'].tab}")
            
            # Ensure the stored mapping data is preserved for future comparisons
            if not hasattr(master_tab, 'stored_mapping_data') and 'final_dataframe' in master_mapping:
                master_tab.stored_mapping_data = master_mapping
            
            # Update the display
            self._update_comparison_display(master_tab, merged_df)
            
            self._update_status(f"Successfully added {new_offer_name} to comparison dataset")
            messagebox.showinfo("Success", f"BOQ '{new_offer_name}' has been added to the comparison dataset!")
            
        except Exception as e:
            # print(f"[DEBUG] Error in _apply_mapping_and_merge: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to merge comparison data: {str(e)}")
    
    def _build_dataframe_from_mapping(self, file_mapping, master_mapping):
        """Build a DataFrame from the provided file mapping using column_index from each mapping."""
        import pandas as pd

        rows = []
        for sheet in getattr(file_mapping, 'sheets', []):
            if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                continue

            col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]

            # Determine valid rows (fallback to include all if no saved validity info)
            validity_dict = {}
            for rc in getattr(sheet, 'row_classifications', []):
                rt = getattr(rc, 'row_type', None)
                if isinstance(rt, str):
                    is_boq = rt.upper() in ['BOQ_ITEM', 'PRIMARY_LINE_ITEM']
                else:
                    is_boq = getattr(rt, 'name', '').upper() in ['BOQ_ITEM', 'PRIMARY_LINE_ITEM']
                validity_dict[rc.row_index] = is_boq

            for rc in getattr(sheet, 'row_classifications', []):
                if not validity_dict.get(rc.row_index, True):
                    continue

                row_data = getattr(rc, 'row_data', None)
                if row_data is None and hasattr(sheet, 'sheet_data'):
                    try:
                        row_data = sheet.sheet_data[rc.row_index]
                    except Exception:
                        row_data = []

                if row_data is None:
                    row_data = []

                row_dict = {}
                for cm in sheet.column_mappings:
                    mt = cm.mapped_type
                    idx = cm.column_index
                    row_dict[mt] = row_data[idx] if idx < len(row_data) else ''

                for mt in col_headers:
                    row_dict.setdefault(mt, '')

                row_dict['sheet'] = sheet.sheet_name
                row_dict['Position'] = getattr(rc, 'position', None)
                rows.append(row_dict)

        return pd.DataFrame(rows) if rows else pd.DataFrame()
    
    def _apply_categories_from_mapping(self, df, master_mapping):
        """Apply categories from master mapping to the new DataFrame"""
        if 'final_dataframe' in master_mapping and 'description' in df.columns:
            saved_df = master_mapping['final_dataframe']
            
            # Handle both possible column names for description
            desc_col = 'description'
            if desc_col not in saved_df.columns and 'Description' in saved_df.columns:
                desc_col = 'Description'
            
            if desc_col in saved_df.columns and 'category' in saved_df.columns:
                category_mapping = {}
                for _, row in saved_df.iterrows():
                    desc = str(row[desc_col]).strip().lower()
                    category = str(row['category']).strip()
                    if desc and category:
                        category_mapping[desc] = category
                
                # Apply categories using the consistent column name
                df['category'] = df['description'].apply(
                    lambda x: category_mapping.get(str(x).strip().lower(), '')
                )
                
                # print(f"[DEBUG] Applied categories to {len(df)} rows, {(df['category'] != '').sum()} matched")
        
        return df
    
    def _merge_comparison_datasets(self, master_df, new_df, new_offer_name):
        """Merge the new dataset with the master comparison dataset"""
        import pandas as pd
        
        # print(f"[DEBUG] Starting merge - Master DF shape: {master_df.shape}, New DF shape: {new_df.shape}")
        
        # CRITICAL FIX: Create copies to avoid modifying the original DataFrames
        master_df_copy = master_df.copy()
        new_df_copy = new_df.copy()
        
        # ROW ORDER PRESERVATION: Store original order from master DataFrame
        master_df_copy['_original_order'] = range(len(master_df_copy))
        # print(f"[DEBUG] Row order preservation: Stored original order for {len(master_df_copy)} rows")
        
        # Reset indices to ensure clean merging
        master_df_copy = master_df_copy.reset_index(drop=True)
        new_df_copy = new_df_copy.reset_index(drop=True)
        
        # print(f"[DEBUG] After reset - Master DF shape: {master_df_copy.shape}, New DF shape: {new_df_copy.shape}")
        
        # Create composite keys for exact matching - use ALL identifying columns
        def create_key(df):
            # Use ALL the identifying columns to create a unique key
            key_parts = []
            for col in ['sheet', 'category', 'code', 'description', 'unit']:
                if col in df.columns:
                    # Clean and normalize the values WITHOUT modifying the original DataFrame
                    values = df[col].fillna('').astype(str).str.strip()
                    key_parts.append(values)
                    # Debug: Check if unit column has values
                    if col == 'unit':
                        unique_units = values.unique()
                        # print(f"[DEBUG] Unit values in key creation: {unique_units[:10]}...")  # Show first 10
                else:
                    key_parts.append(pd.Series([''] * len(df)))
            
            # Combine all parts with a separator
            return key_parts[0] + '|' + key_parts[1] + '|' + key_parts[2] + '|' + key_parts[3] + '|' + key_parts[4]
        
        # Debug: Show unique sheet names in both datasets
        # print(f"[DEBUG] Master dataset sheets: {sorted(master_df_copy['sheet'].unique())}")
        # print(f"[DEBUG] New dataset sheets: {sorted(new_df_copy['sheet'].unique())}")
        
        # Check if sheet names match
        master_sheets = set(master_df_copy['sheet'].unique())
        new_sheets = set(new_df_copy['sheet'].unique())
        if master_sheets != new_sheets:
            # print(f"[DEBUG] WARNING: Sheet names don't match!")
            # print(f"[DEBUG] Sheets only in master: {master_sheets - new_sheets}")
            # print(f"[DEBUG] Sheets only in new: {new_sheets - master_sheets}")
            pass
        else:
            # print(f"[DEBUG] Sheet names match perfectly!")
            pass
        
        # Add keys to the COPIES, not the originals
        master_df_copy['_key'] = create_key(master_df_copy)
        new_df_copy['_key'] = create_key(new_df_copy)
        
        # print(f"[DEBUG] Master unique keys: {master_df_copy['_key'].nunique()}")
        # print(f"[DEBUG] New unique keys: {new_df_copy['_key'].nunique()}")
        # print(f"[DEBUG] Common keys: {len(set(master_df_copy['_key']) & set(new_df_copy['_key']))}")
        
        # CRITICAL: Check for duplicate keys within each dataset
        master_duplicates = master_df_copy[master_df_copy['_key'].duplicated(keep=False)]
        new_duplicates = new_df_copy[new_df_copy['_key'].duplicated(keep=False)]
        
        if len(master_duplicates) > 0:
            # print(f"[DEBUG] WARNING: Found {len(master_duplicates)} duplicate keys in MASTER dataset!")
            duplicate_keys = master_duplicates['_key'].unique()
            for dup_key in duplicate_keys[:3]:  # Show first 3 examples
                dup_rows = master_df_copy[master_df_copy['_key'] == dup_key]
                # print(f"[DEBUG] Master duplicate key '{dup_key}' appears {len(dup_rows)} times:")
                for idx, row in dup_rows.iterrows():
                    # print(f"[DEBUG]   Row {idx}: sheet='{row.get('sheet', '')}', code='{row.get('code', '')}', desc='{str(row.get('description', ''))[:50]}...'")
        
                    pass
        if len(new_duplicates) > 0:
            # print(f"[DEBUG] WARNING: Found {len(new_duplicates)} duplicate keys in NEW dataset!")
            duplicate_keys = new_duplicates['_key'].unique()
            for dup_key in duplicate_keys[:3]:  # Show first 3 examples
                dup_rows = new_df_copy[new_df_copy['_key'] == dup_key]
                # print(f"[DEBUG] New duplicate key '{dup_key}' appears {len(dup_rows)} times:")
                for idx, row in dup_rows.iterrows():
                    # print(f"[DEBUG]   Row {idx}: sheet='{row.get('sheet', '')}', code='{row.get('code', '')}', desc='{str(row.get('description', ''))[:50]}...'")
        
        # Handle duplicate keys by using row position within each unique key group
                    pass
        if not master_duplicates.empty or not new_duplicates.empty:
            # print(f"[DEBUG] Making keys unique by adding row position within duplicate groups...")
            
            # Use a robust method to create unique keys using groupby and cumcount
            def create_unique_keys_with_position(df):
                # ROW ORDER PRESERVATION: Do NOT sort by key - preserve original order
                # Instead, use Position field or original order to maintain Excel row sequence
                df_with_pos = df.copy()
                
                # Create a positional counter within each group of duplicate keys
                df_with_pos['_pos'] = df_with_pos.groupby('_key').cumcount()
                
                # Create the unique key by combining the base key and the position
                df_with_pos['_unique_key'] = df_with_pos['_key'] + '|POS_' + df_with_pos['_pos'].astype(str)
                
                # ROW ORDER PRESERVATION: Sort by Position field to maintain Excel row order
                if 'Position' in df_with_pos.columns:
                    # Extract row number from Position field (format: sheet_name_row_number)
                    try:
                        df_with_pos['_row_num'] = df_with_pos['Position'].str.extract(r'_(\d+)$')[0].astype(int)
                        df_with_pos = df_with_pos.sort_values('_row_num').reset_index(drop=True)
                        df_with_pos = df_with_pos.drop(columns=['_row_num'])
                        # print(f"[DEBUG] Row order preservation: Sorted by Position field to maintain Excel row order")
                    except Exception as e:
                        # print(f"[DEBUG] Row order preservation: Could not sort by Position field: {e}")
                        # Fallback: preserve original order by not sorting
                        pass
                elif '_original_order' in df_with_pos.columns:
                    # Fallback: sort by original order
                    df_with_pos = df_with_pos.sort_values('_original_order').reset_index(drop=True)
                    # print(f"[DEBUG] Row order preservation: Sorted by original order")
                else:
                    # print(f"[DEBUG] Row order preservation: No Position or original order field found, preserving current order")
                
                # Drop the temporary position column
                    pass
                df_with_pos = df_with_pos.drop(columns=['_pos'])
                return df_with_pos

            # Apply the function to both DataFrames
            master_df_copy = create_unique_keys_with_position(master_df_copy)
            new_df_copy = create_unique_keys_with_position(new_df_copy)
            
            key_column = '_unique_key'
        else:
            # No duplicates, use original keys
            master_df_copy['_unique_key'] = master_df_copy['_key']
            new_df_copy['_unique_key'] = new_df_copy['_key']
            key_column = '_unique_key'
        
        # Prepare new offer columns
        new_columns = {
            'quantity': f'quantity[{new_offer_name}]',
            'unit_price': f'unit_price[{new_offer_name}]',
            'total_price': f'total_price[{new_offer_name}]',
            'manhours': f'manhours[{new_offer_name}]',
            'wage': f'wage[{new_offer_name}]'
        }
        
        # Create a mapping from key to new offer values
        new_offer_mapping = {}
        # print(f"[DEBUG] New DataFrame columns: {new_df_copy.columns.tolist()}")
        # print(f"[DEBUG] Looking for columns: {list(new_columns.keys())}")
        
        for idx, row in new_df_copy.iterrows():
            key = row[key_column]  # Use the unique key
            values = {}
            for old_col, new_col in new_columns.items():
                if old_col in new_df_copy.columns:
                    value = row[old_col]
                    values[new_col] = value
                    # Debug first few rows
                    if idx < 3:
                        # print(f"[DEBUG] Row {idx}: {old_col} = {value} (type: {type(value)})")
                        pass
                else:
                    values[new_col] = 0
                    if idx < 3:
                        # print(f"[DEBUG] Row {idx}: {old_col} NOT FOUND, using 0")
                        pass
            new_offer_mapping[key] = values
        
        # print(f"[DEBUG] Created mapping for {len(new_offer_mapping)} unique keys")
        
        # Debug: Show a sample of the mapping
        sample_keys = list(new_offer_mapping.keys())[:3]
        for sample_key in sample_keys:
            # print(f"[DEBUG] Sample mapping '{sample_key}': {new_offer_mapping[sample_key]}")
        
        # Check for exact match in structure using unique keys
            pass
        master_unique_keys = set(master_df_copy[key_column])
        new_unique_keys = set(new_df_copy[key_column])
        
        if master_unique_keys != new_unique_keys:
            # print(f"[DEBUG] WARNING: Unique key sets don't match exactly!")
            # print(f"[DEBUG] Keys only in master: {len(master_unique_keys - new_unique_keys)}")
            # print(f"[DEBUG] Keys only in new: {len(new_unique_keys - master_unique_keys)}")
            
            # Show some examples of mismatched keys
            only_in_master = list(master_unique_keys - new_unique_keys)[:3]
            only_in_new = list(new_unique_keys - master_unique_keys)[:3]
            # print(f"[DEBUG] Sample keys only in master: {only_in_master}")
            # print(f"[DEBUG] Sample keys only in new: {only_in_new}")
        else:
            # print(f"[DEBUG] PERFECT MATCH: All unique keys match exactly! This is expected for the same offer.")
        
        # Start with the master DataFrame structure
            pass
        merged_df = master_df_copy.copy()
        
        # Debug: Check if unit column exists and has values before merge
        if 'unit' in merged_df.columns:
            unit_values_before = merged_df['unit'].value_counts()
            # print(f"[DEBUG] Unit column before merge - unique values: {unit_values_before.to_dict()}")
        else:
            # print(f"[DEBUG] WARNING: Unit column not found in master DataFrame!")
        
        # CRITICAL FIX: Preserve unit column values during merge
        # Store the unit column values from the ORIGINAL master DataFrame (not the processed copy)
            pass
        unit_backup = None
        if 'unit' in master_df.columns:
            unit_backup = master_df['unit'].copy()
            # print(f"[DEBUG] Backed up unit column from ORIGINAL master DataFrame with {len(unit_backup)} values")
            original_unit_values = master_df['unit'].value_counts()
            # print(f"[DEBUG] Original unit values: {original_unit_values.to_dict()}")
            
            # Check for invalid unit values in the original master DataFrame
            invalid_master_units = master_df[master_df['unit'].isin(['Quantity', 'quantity', 'Description', 'description', 'Code', 'code'])]
            if len(invalid_master_units) > 0:
                # print(f"[DEBUG] ⚠️  CRITICAL: Original master DataFrame contains {len(invalid_master_units)} INVALID unit values!")
                # print(f"[DEBUG] Sample invalid master unit rows:")
                for idx, row in invalid_master_units.head(3).iterrows():
                    # print(f"[DEBUG]   Row {idx}: unit='{row['unit']}', code='{row.get('code', 'N/A')}', description='{row.get('description', 'N/A')[:50]}...'")
                    pass
        elif 'unit' in merged_df.columns:
            unit_backup = merged_df['unit'].copy()
            # print(f"[DEBUG] Backed up unit column from processed copy with {len(unit_backup)} values")
        
        # Add the new offer columns
        for new_col in new_columns.values():
            merged_df[new_col] = 0  # Initialize with zeros
        
        # Fill in the values from the new offer using the unique key mapping
        matched_count = 0
        for idx, row in merged_df.iterrows():
            key = row[key_column]  # Use the unique key
            if key in new_offer_mapping:
                for new_col, value in new_offer_mapping[key].items():
                    merged_df.at[idx, new_col] = value
                    # Debug first few assignments
                    if idx < 3:
                        # print(f"[DEBUG] Assigned row {idx}, col {new_col} = {value}")
                        pass
                matched_count += 1
            else:
                # If the key from master is not in the new offer mapping, it means this row is missing in the new offer
                # We should fill its new offer columns with 0
                for new_col in new_columns.values():
                    merged_df.at[idx, new_col] = 0
                if idx < 3:
                    # print(f"[DEBUG] Row {idx}: Key '{key}' not found in mapping, filling with 0")
        
        # print(f"[DEBUG] Successfully matched {matched_count} out of {len(merged_df)} rows")
        
        # Debug: Check the actual values being assigned
                    pass
        total_assigned_quantity = 0
        total_assigned_unit_price = 0
        total_assigned_total_price = 0
        total_assigned_manhours = 0
        total_assigned_wage = 0
        
        for new_col in new_columns.values():
            if 'quantity' in new_col:
                col_sum = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0).sum()
                total_assigned_quantity += col_sum
                # print(f"[DEBUG] Total {new_col}: {col_sum}")
            elif 'unit_price' in new_col:
                col_sum = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0).sum()
                total_assigned_unit_price += col_sum
                # print(f"[DEBUG] Total {new_col}: {col_sum}")
            elif 'total_price' in new_col:
                col_sum = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0).sum()
                total_assigned_total_price += col_sum
                # print(f"[DEBUG] Total {new_col}: {col_sum}")
            elif 'manhours' in new_col:
                col_sum = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0).sum()
                total_assigned_manhours += col_sum
                # print(f"[DEBUG] Total {new_col}: {col_sum}")
            elif 'wage' in new_col:
                col_sum = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0).sum()
                total_assigned_wage += col_sum
                # print(f"[DEBUG] Total {new_col}: {col_sum}")
        
        # print(f"[DEBUG] Total assigned - Quantity: {total_assigned_quantity}, Unit Price: {total_assigned_unit_price}, Total Price: {total_assigned_total_price}, Manhours: {total_assigned_manhours}, Wage: {total_assigned_wage}")
        
        # Handle any new rows that don't exist in master (shouldn't happen in comparison mode, but just in case)
        unmatched_keys = set(new_offer_mapping.keys()) - set(merged_df[key_column])
        if unmatched_keys:
            # print(f"[DEBUG] Found {len(unmatched_keys)} unmatched keys from new dataset, adding them to the master set.")
            
            # Create a list to hold the new rows
            rows_to_add = []
            
            # Get the columns from the master DataFrame to ensure consistency
            master_cols = merged_df.columns
            
            # Add unmatched rows
            for key in unmatched_keys:
                # Find the original row in new_df_copy
                new_row_series = new_df_copy[new_df_copy[key_column] == key].iloc[0]
                
                # Create a new row for the merged dataset
                merged_row = {col: None for col in master_cols} # Initialize with None
                
                # Copy base columns
                for col in ['sheet', 'category', 'code', 'description', 'unit']:
                    if col in new_row_series:
                        merged_row[col] = new_row_series[col]
            
                # Initialize all existing offer columns with 0
                for col in master_cols:
                    if col.startswith(('quantity[', 'unit_price[', 'total_price[')):
                        merged_row[col] = 0
            
                # Set the new offer values
                for new_col, value in new_offer_mapping[key].items():
                    merged_row[new_col] = value
            
                # Add the unique key temporarily for debugging
                merged_row[key_column] = key
            
                # Append to list
                rows_to_add.append(merged_row)
            
            # Convert list of dicts to DataFrame and concatenate
            if rows_to_add:
                unmatched_df = pd.DataFrame(rows_to_add)
                merged_df = pd.concat([merged_df, unmatched_df], ignore_index=True)
                # print(f"[DEBUG] Added {len(unmatched_df)} new rows to the comparison.")
        
        # Remove the temporary key columns
        columns_to_drop = ['_key', '_unique_key']
        for col in columns_to_drop:
            if col in merged_df.columns:
                merged_df = merged_df.drop(col, axis=1)
        
        # Reorder columns to maintain the specified order
        base_columns = ['sheet', 'category', 'code', 'description', 'unit']
        
        # Get all offer columns in the order they were added
        offer_columns = []
        for col in merged_df.columns:
            if col.startswith(('quantity[', 'unit_price[', 'total_price[', 'manhours[', 'wage[')):
                offer_columns.append(col)
        
        # Sort offer columns by offer name to maintain consistent order
        offer_names = set()
        for col in offer_columns:
            if '[' in col and ']' in col:
                offer_name = col.split('[')[1].split(']')[0]
                offer_names.add(offer_name)
        
        offer_names = sorted(offer_names)
        ordered_offer_columns = []
        for offer in offer_names:
            for prefix in ['quantity', 'unit_price', 'total_price', 'manhours', 'wage']:
                col_name = f'{prefix}[{offer}]'
                if col_name in merged_df.columns:
                    ordered_offer_columns.append(col_name)
        
        # Add scope if it exists (it doesn't get offer-specific suffixes)
        if 'scope' in merged_df.columns and 'scope' not in base_columns:
            base_columns.append('scope')
        
        # Final column order
        final_columns = base_columns + ordered_offer_columns
        
        # Only include columns that exist
        final_columns = [col for col in final_columns if col in merged_df.columns]
        merged_df = merged_df[final_columns]
        
        # CRITICAL FIX: Restore unit column values if they were lost during merge
        if unit_backup is not None and 'unit' in merged_df.columns:
            # Check if unit column was corrupted during merge
            unit_values_after = merged_df['unit'].value_counts()
            # print(f"[DEBUG] Unit column after merge - unique values: {unit_values_after.to_dict()}")
            
            # If unit column is mostly empty but we have backup, restore it
            empty_count = (merged_df['unit'] == '').sum() + merged_df['unit'].isna().sum()
            total_count = len(merged_df)
            
            if empty_count > total_count * 0.5:  # If more than 50% empty, restore from backup
                # print(f"[DEBUG] Unit column corrupted ({empty_count}/{total_count} empty), restoring from backup")
                merged_df['unit'] = unit_backup
                unit_values_restored = merged_df['unit'].value_counts()
                # print(f"[DEBUG] Unit column restored - unique values: {unit_values_restored.to_dict()}")
            else:
                # print(f"[DEBUG] Unit column preserved correctly ({empty_count}/{total_count} empty)")
        
        # Final debug of unit column
                pass
        if 'unit' in merged_df.columns:
            final_unit_values = merged_df['unit'].value_counts()
            # print(f"[DEBUG] FINAL unit column values: {final_unit_values.to_dict()}")
        else:
            # print(f"[DEBUG] FINAL: Unit column missing from merged DataFrame!")
        
        # Debug: Check if unit column still exists and has values after reordering
            pass
        if 'unit' in merged_df.columns:
            unit_values_after = merged_df['unit'].value_counts()
            # print(f"[DEBUG] Unit column after merge - unique values: {unit_values_after.to_dict()}")
            
            # Check for invalid unit values that suggest data corruption
            invalid_units = merged_df[merged_df['unit'].isin(['Quantity', 'quantity', 'Description', 'description', 'Code', 'code'])]
            if len(invalid_units) > 0:
                # print(f"[DEBUG] ⚠️  CRITICAL: Found {len(invalid_units)} rows with INVALID unit values!")
                # print(f"[DEBUG] Sample invalid unit rows:")
                for idx, row in invalid_units.head(3).iterrows():
                    # print(f"[DEBUG]   Row {idx}: unit='{row['unit']}', code='{row.get('code', 'N/A')}', description='{row.get('description', 'N/A')[:50]}...'")
                    
            # Check if any unit values are empty when they shouldn't be
                    pass
            empty_units = (merged_df['unit'] == '') | (merged_df['unit'].isna())
            if empty_units.any():
                # print(f"[DEBUG] WARNING: Found {empty_units.sum()} empty unit values after merge!")
                pass
        else:
            # print(f"[DEBUG] CRITICAL: Unit column missing after column reordering!")
        
        # Skip duplicate removal - we already handled duplicates properly during key generation
        # Removing duplicates here would corrupt the data since identical rows in comparison
        # scenarios (like comparing the same file twice) would be incorrectly removed
        # print(f"[DEBUG] Skipping duplicate removal to preserve data integrity")
        
        # Verify the merge worked correctly
        # print(f"[DEBUG] Final merged dataset has {len(merged_df)} rows and columns: {merged_df.columns.tolist()}")
        
        # Check that the new offer columns have reasonable values
            pass
        for new_col in new_columns.values():
            if new_col in merged_df.columns:
                non_zero_count = (merged_df[new_col] != 0).sum()
                # Convert to numeric and handle errors safely
                try:
                    numeric_col = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0)
                    total_sum = numeric_col.sum()
                    # print(f"[DEBUG] Column {new_col}: {non_zero_count} non-zero values, sum = {total_sum}")
                except Exception as e:
                    # print(f"[DEBUG] Column {new_col}: {non_zero_count} non-zero values, sum calculation failed: {e}")
                    # Fix the column to be numeric
                    merged_df[new_col] = pd.to_numeric(merged_df[new_col], errors='coerce').fillna(0)
        
        # ROW ORDER PRESERVATION: Sort by original order to maintain first file's row sequence
        if '_original_order' in merged_df.columns:
            # print(f"[DEBUG] Row order preservation: Sorting final DataFrame by original order")
            merged_df = merged_df.sort_values('_original_order').reset_index(drop=True)
            # Remove the temporary ordering column
            merged_df = merged_df.drop(columns=['_original_order'])
            # print(f"[DEBUG] Row order preservation: Final DataFrame sorted and cleaned")
        elif 'Position' in merged_df.columns:
            # Fallback: sort by Position field to maintain Excel row order
            try:
                # print(f"[DEBUG] Row order preservation: Sorting by Position field as fallback")
                merged_df['_row_num'] = merged_df['Position'].str.extract(r'_(\d+)$')[0].astype(int)
                merged_df = merged_df.sort_values('_row_num').reset_index(drop=True)
                merged_df = merged_df.drop(columns=['_row_num'])
                # print(f"[DEBUG] Row order preservation: Final DataFrame sorted by Position field")
            except Exception as e:
                # print(f"[DEBUG] Row order preservation: Could not sort by Position field: {e}")
                pass
        else:
            # print(f"[DEBUG] Row order preservation: No ordering fields found, preserving current order")
        
        # Sample of merged data for verification
        # print(f"[DEBUG] Sample of merged data:")
            pass
        print(merged_df.head(3).to_string())
        
        return merged_df

    def _on_confirm_mapped_file_review(self, file_mapping):
        """Simplified confirmation handler for mapped-file row review, mirroring the normal workflow."""
        self._update_status("Row review confirmed. Starting categorization process...")

        import pandas as pd
        rows = []

        for sheet in getattr(file_mapping, 'sheets', []):
            if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                continue

            # Build list of mapped column types in order
            col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]
            validity_dict = self.row_validity.get(sheet.sheet_name, {})

            for rc in getattr(sheet, 'row_classifications', []):
                # Include only rows marked as valid/BOQ items
                if not validity_dict.get(rc.row_index, True):
                    continue

                row_data = getattr(rc, 'row_data', None)
                if row_data is None and hasattr(sheet, 'sheet_data'):
                    try:
                        row_data = sheet.sheet_data[rc.row_index]
                    except Exception:
                        row_data = []

                if row_data is None:
                    row_data = []

                # Build a dict using the *column_index* information from each ColumnMapping
                row_dict = {}
                for cm in sheet.column_mappings:
                    mapped_type = cm.mapped_type
                    idx = cm.column_index
                    row_dict[mapped_type] = row_data[idx] if idx < len(row_data) else ''

                # Ensure any missing mapped types are present (empty string)
                for mt in col_headers:
                    row_dict.setdefault(mt, '')

                row_dict['Source_Sheet'] = sheet.sheet_name
                row_dict['Position'] = getattr(rc, 'position', None)
                rows.append(row_dict)

        # Convert to DataFrame
        if rows:
            df = pd.DataFrame(rows)
            if 'description' in df.columns and 'Description' not in df.columns:
                df.rename(columns={'description': 'Description'}, inplace=True)
        else:
            df = pd.DataFrame()

        file_mapping.dataframe = df

        if df.empty or 'Description' not in df.columns:
            messagebox.showerror("Categorization Error", "No valid data found for categorization.")
            return

        # Proceed to categorization dialog
        self._start_categorization(file_mapping)

    def _validate_position_data_integrity(self, mapping_data):
        """
        Validate that position data is present and properly formatted in mapping data
        
        Args:
            mapping_data: Dictionary containing mapping information
            
        Returns:
            bool: True if position data is valid, False otherwise
        """
        if not mapping_data or 'sheets' not in mapping_data:
            return True  # No data to validate
        
        for sheet in mapping_data['sheets']:
            sheet_name = sheet.get('sheet_name', 'Unknown')
            row_classifications = sheet.get('row_classifications', [])
            
            for i, rc in enumerate(row_classifications):
                # Check if position data exists
                if 'position' not in rc:
                    print(f"    Missing position data for row {i} in sheet {sheet_name}")
                    # Generate position if missing
                    rc['position'] = f"{sheet_name}_{i + 1}"
                    continue
                
                position = rc['position']
                
                # Validate position format (should be sheet_name_row_number)
                if not isinstance(position, str) or '_' not in position:
                    print(f"    Fixed invalid position format for row {i}: '{position}'")
                    rc['position'] = f"{sheet_name}_{i + 1}"
                    continue
                
                # Additional validation: check if it follows the expected pattern
                parts = position.split('_')
                if len(parts) < 2 or not parts[-1].isdigit():
                    print(f"    Fixed invalid position format for row {i}: '{position}'")
                    rc['position'] = f"{sheet_name}_{i + 1}"
        
        return True

    def _validate_position_description_match(self, current_file_mapping, saved_mapping_data):
        """
        Validate that position-description pairs match exactly between current file and saved mapping
        
        Args:
            current_file_mapping: Current file mapping object
            saved_mapping_data: Saved mapping data dictionary
            
        Returns:
            dict: Validation result with detailed information
                {
                    'is_valid': bool,
                    'errors': list of error messages,
                    'mismatched_positions': list of position details,
                    'summary': summary string
                }
        """
        validation_result = {
            'is_valid': True,
            'errors': [],
            'mismatched_positions': [],
            'summary': ''
        }
        
        logger.debug("Starting position-description validation")
        
        try:
            # Get saved mapping data
            saved_sheets = saved_mapping_data.get('sheets', [])
            saved_position_descriptions = {}
            
            # Extract position-description pairs from saved mapping
            for saved_sheet in saved_sheets:
                sheet_name = saved_sheet.get('sheet_name', '')
                row_classifications = saved_sheet.get('row_classifications', [])
                
                for rc in row_classifications:
                    if isinstance(rc, dict):
                        position = rc.get('position', None)
                        row_data = rc.get('row_data', [])
                        
                        # Get description (usually first column)
                        description = ''
                        if row_data and len(row_data) > 0:
                            description = str(row_data[0]).strip()
                        
                        if position and description:
                            saved_position_descriptions[position] = {
                                'description': description,
                                'sheet_name': sheet_name,
                                'row_data': row_data
                            }
            
            logger.debug(f"Found {len(saved_position_descriptions)} position-description pairs in saved mapping")
            
            # Extract position-description pairs from current file
            current_position_descriptions = {}
            
            for sheet in getattr(current_file_mapping, 'sheets', []):
                sheet_name = sheet.sheet_name
                
                if hasattr(sheet, 'row_classifications'):
                    for rc in sheet.row_classifications:
                        position = getattr(rc, 'position', None)
                        row_data = getattr(rc, 'row_data', None)
                        
                        # Get row data if not available
                        if row_data is None and hasattr(sheet, 'sheet_data'):
                            try:
                                row_data = sheet.sheet_data[rc.row_index]
                            except Exception:
                                row_data = None
                        
                        if row_data is None:
                            row_data = []
                        
                        # Get description (usually first column)
                        description = ''
                        if row_data and len(row_data) > 0:
                            description = str(row_data[0]).strip()
                        
                        if position and description:
                            current_position_descriptions[position] = {
                                'description': description,
                                'sheet_name': sheet_name,
                                'row_data': row_data
                            }
            
            logger.debug(f"Found {len(current_position_descriptions)} position-description pairs in current file")
            
            # Compare position-description pairs
            mismatched_count = 0
            missing_positions = []
            description_mismatches = []
            
            # Check that all saved positions exist in current file
            for saved_position, saved_data in saved_position_descriptions.items():
                if saved_position not in current_position_descriptions:
                    missing_positions.append({
                        'position': saved_position,
                        'expected_description': saved_data['description'],
                        'sheet_name': saved_data['sheet_name']
                    })
                    mismatched_count += 1
                else:
                    # Check if descriptions match
                    current_data = current_position_descriptions[saved_position]
                    saved_desc = saved_data['description']
                    current_desc = current_data['description']
                    
                    if saved_desc != current_desc:
                        description_mismatches.append({
                            'position': saved_position,
                            'sheet_name': saved_data['sheet_name'],
                            'expected_description': saved_desc,
                            'actual_description': current_desc
                        })
                        mismatched_count += 1
            
            # Check for extra positions in current file (not in saved mapping)
            extra_positions = []
            for current_position in current_position_descriptions:
                if current_position not in saved_position_descriptions:
                    current_data = current_position_descriptions[current_position]
                    extra_positions.append({
                        'position': current_position,
                        'description': current_data['description'],
                        'sheet_name': current_data['sheet_name']
                    })
            
            # Log detailed validation results
            logger.debug(f"Validation analysis: {len(missing_positions)} missing positions, "
                        f"{len(description_mismatches)} description mismatches, "
                        f"{len(extra_positions)} extra positions")
            
            # Generate validation result
            if missing_positions:
                validation_result['is_valid'] = False
                logger.warning(f"Found {len(missing_positions)} missing positions")
                for missing in missing_positions:
                    error_msg = f"Missing position {missing['position']} in sheet '{missing['sheet_name']}' - expected description: '{missing['expected_description']}'"
                    validation_result['errors'].append(error_msg)
                    validation_result['mismatched_positions'].append(missing)
                    logger.debug(f"Missing position details: {missing}")
            
            if description_mismatches:
                validation_result['is_valid'] = False
                logger.warning(f"Found {len(description_mismatches)} description mismatches")
                for mismatch in description_mismatches:
                    error_msg = f"Description mismatch at position {mismatch['position']} in sheet '{mismatch['sheet_name']}' - expected: '{mismatch['expected_description']}', actual: '{mismatch['actual_description']}'"
                    validation_result['errors'].append(error_msg)
                    validation_result['mismatched_positions'].append(mismatch)
                    logger.debug(f"Description mismatch details: {mismatch}")
            
            if extra_positions:
                # Extra positions are not necessarily an error, but we'll log them
                logger.info(f"Found {len(extra_positions)} extra positions in current file (not in saved mapping)")
                for extra in extra_positions:
                    warning_msg = f"Extra position {extra['position']} in sheet '{extra['sheet_name']}' with description: '{extra['description']}'"
                    logger.debug(warning_msg)
            
            # Generate summary
            if validation_result['is_valid']:
                validation_result['summary'] = f"Position-description validation passed. {len(saved_position_descriptions)} positions matched successfully."
                logger.info(validation_result['summary'])
            else:
                validation_result['summary'] = f"Position-description validation failed. {mismatched_count} mismatches found: {len(missing_positions)} missing positions, {len(description_mismatches)} description mismatches."
                logger.error(validation_result['summary'])
            
            return validation_result
            
        except Exception as e:
            validation_result['is_valid'] = False
            validation_result['errors'].append(f"Error during position-description validation: {str(e)}")
            validation_result['summary'] = f"Validation failed due to error: {str(e)}"
            logger.error(f"Position-description validation error: {e}", exc_info=True)
            return validation_result

    def _collect_summary_data(self, tab=None):
        """Collect summary data for the new summary grid, using the current tab's DataFrame if available (like the Summarize grid)."""
        summary_data = []
        
        # Use the current tab's DataFrame if available
        if tab is not None and hasattr(tab, 'final_dataframe') and tab.final_dataframe is not None:
            df = tab.final_dataframe
            file_data = None
            # Try to find the file_data for this tab (for offer info)
            for fd in self.controller.current_files.values():
                if hasattr(fd['file_mapping'], 'tab') and fd['file_mapping'].tab == tab:
                    file_data = fd
                    break
            
            # If not found by tab reference, try to find by checking if this tab has a final_dataframe
            if file_data is None:
                for fd in self.controller.current_files.values():
                    if 'file_mapping' in fd and hasattr(fd['file_mapping'], 'categorized_dataframe'):
                        # Check if this file has the same DataFrame as the current tab
                        if hasattr(tab, 'final_dataframe') and tab.final_dataframe is not None:
                            if fd['file_mapping'].categorized_dataframe is tab.final_dataframe:
                                file_data = fd
                                break
            
            print(f"[DEBUG] Found file_data for tab: {file_data is not None}")
            if file_data:
                print(f"[DEBUG] file_data keys: {list(file_data.keys())}")
                if 'offers' in file_data:
                    print(f"[DEBUG] file_data offers: {list(file_data['offers'].keys())}")
            else:
                print(f"[DEBUG] No file_data found for tab. Available files: {list(self.controller.current_files.keys())}")
                for fd in self.controller.current_files.values():
                    if 'file_mapping' in fd:
                        tab_ref = getattr(fd['file_mapping'], 'tab', None)
                        print(f"[DEBUG] File mapping tab: {tab_ref}, current tab: {tab}")
                
                # DEBUG: Show what's actually stored in controller.current_files
                print(f"[DEBUG] === DETAILED FILE DATA INSPECTION ===")
                for file_key, fd in self.controller.current_files.items():
                    print(f"[DEBUG] File: {file_key}")
                    print(f"[DEBUG]   Keys: {list(fd.keys())}")
                    if 'offers' in fd:
                        print(f"[DEBUG]   Offers: {list(fd['offers'].keys())}")
                        for offer_key, offer_info in fd['offers'].items():
                            print(f"[DEBUG]     {offer_key}: {offer_info}")
                    if 'file_mapping' in fd:
                        tab_ref = getattr(fd['file_mapping'], 'tab', None)
                        print(f"[DEBUG]   Tab reference: {tab_ref}")
                print(f"[DEBUG] === END INSPECTION ===")
                
                # Try to find file_data by checking if any file has offers that match our offer keys
                for fd in self.controller.current_files.values():
                    if 'offers' in fd:
                        print(f"[DEBUG] Found file with offers: {list(fd['offers'].keys())}")
                        # Use the first file with offers as our file_data
                        file_data = fd
                        print(f"[DEBUG] Using fallback file_data with offers: {list(file_data['offers'].keys())}")
                        break
            
            if df is not None:
                print(f"[DEBUG] Using tab.final_dataframe with columns: {list(df.columns)}")
                
                # Robust check for comparison columns
                def is_comparison_col(col):
                    col_clean = col.replace(' ', '').lower()
                    # Only consider it a comparison column if it has total_price AND contains brackets or underscores
                    # This excludes plain 'total_price' columns which are single-offer
                    return (col_clean.startswith('total_price[') and ']' in col_clean) or \
                           (col_clean.startswith('total_price_') and '_' in col_clean[12:])  # After 'total_price_'
                
                if any(is_comparison_col(col) for col in df.columns):
                    # Comparison dataset: multiple offers
                    offer_columns = {}
                    for col in df.columns:
                        col_clean = col.replace(' ', '').lower()
                        if col_clean.startswith('total_price[') and ']' in col:
                            offer_key = col[col.find('[')+1:col.find(']')]
                            offer_columns[offer_key] = col
                        elif col_clean.startswith('total_price_'):
                            offer_key = col.split('_', 1)[1]
                            offer_columns[offer_key] = col
                    
                    offers_info = file_data.get('offers', {}) if file_data else {}
                    print(f"[DEBUG] Comparison dataset - offers found: {list(offer_columns.keys())}")
                    
                    for offer_key, price_col in offer_columns.items():
                        offer_info = offers_info.get(offer_key, {})
                        supplier = offer_info.get('supplier_name', offer_key)
                        project_name = offer_info.get('project_name', 'Unknown')
                        date = offer_info.get('date', 'Unknown')
                        
                        print(f"[DEBUG] Looking up offer {offer_key}: supplier={supplier}, project={project_name}, date={date}")
                        print(f"[DEBUG] Available offers in offers_info: {list(offers_info.keys())}")
                        for k, v in offers_info.items():
                            print(f"[DEBUG]   {k}: {v}")
                        
                        # If offer info is not found in file_data, try to get from current_offer_info
                        if supplier == offer_key and project_name == 'Unknown' and date == 'Unknown':
                            current_offer_info = getattr(self, 'current_offer_info', {})
                            if current_offer_info and current_offer_info.get('supplier_name') == offer_key:
                                supplier = current_offer_info.get('supplier_name', supplier)
                                project_name = current_offer_info.get('project_name', project_name)
                                date = current_offer_info.get('date', date)
                                print(f"[DEBUG] Using current_offer_info for {offer_key}: supplier={supplier}, project={project_name}, date={date}")
                        
                        # Calculate total price with proper null checks
                        if file_data and 'file_mapping' in file_data:
                            total_price = self._calculate_total_price(file_data['file_mapping'], offer_name=offer_key)
                        else:
                            # Fallback: calculate from the DataFrame directly
                            try:
                                if price_col in df.columns:
                                    # Convert to numeric and handle non-numeric values
                                    numeric_values = pd.to_numeric(df[price_col], errors='coerce')
                                    total_price = numeric_values.sum()
                                    print(f"[DEBUG] Calculated total for {offer_key} from DataFrame: {total_price}")
                                else:
                                    total_price = 0.0
                            except Exception as e:
                                print(f"[DEBUG] Error calculating total price for {offer_key}: {e}")
                                total_price = 0.0
                        
                        # Ensure total_price is a valid number
                        try:
                            if isinstance(total_price, str):
                                total_price = float(total_price)
                            elif not isinstance(total_price, (int, float)):
                                total_price = 0.0
                        except (ValueError, TypeError):
                            print(f"[DEBUG] Invalid total_price for {offer_key}: {total_price}, setting to 0.0")
                            total_price = 0.0
                        
                        print(f"[DEBUG] Summary row: offer_key={offer_key}, supplier={supplier}, project={project_name}, date={date}, total_price={total_price}")
                        summary_data.append([supplier, project_name, date, total_price])
                else:
                    # Single-offer dataset
                    offer_info = file_data.get('offer_info', {}) if file_data else {}
                    print(f"[DEBUG] Single-offer dataset - file_data offer_info: {offer_info}")
                    
                    # Enhanced fallback logic for single-offer datasets
                    supplier = offer_info.get('supplier_name', 'Unknown')
                    project_name = offer_info.get('project_name', 'Unknown')
                    date = offer_info.get('date', 'Unknown')
                    
                    # If offer_info is empty or supplier is Unknown, try to get from current_offer_info
                    if not offer_info or supplier == 'Unknown':
                        current_offer_info = getattr(self, 'current_offer_info', {})
                        print(f"[DEBUG] Trying current_offer_info fallback: {current_offer_info}")
                        if current_offer_info:
                            supplier = current_offer_info.get('supplier_name', supplier)
                            project_name = current_offer_info.get('project_name', project_name)
                            date = current_offer_info.get('date', date)
                    
                    # If still Unknown, try to get from current_offer_name
                    if supplier == 'Unknown':
                        current_offer_name = getattr(self, 'current_offer_name', None)
                        print(f"[DEBUG] Trying current_offer_name fallback: {current_offer_name}")
                        if current_offer_name:
                            supplier = current_offer_name
                    
                    total_price = self._calculate_total_price(file_data['file_mapping'], offer_name=supplier)
                    print(f"[DEBUG] Summary row (single): supplier={supplier}, project={project_name}, date={date}, total_price={total_price}")
                    summary_data.append([supplier, project_name, date, total_price])
        else:
            # Fallback to original logic for backward compatibility
            for file_key, file_data in self.controller.current_files.items():
                file_mapping = file_data['file_mapping']
                # Try to detect if this is a comparison dataset
                df = None
                if hasattr(file_mapping, 'categorized_dataframe') and file_mapping.categorized_dataframe is not None:
                    df = file_mapping.categorized_dataframe
                else:
                    try:
                        df = self._build_final_grid_dataframe(file_mapping)
                    except Exception:
                        df = None
                
                if df is not None:
                    print(f"[DEBUG] DataFrame columns for file {file_key}: {list(df.columns)}")
                    
                    # Robust check for comparison columns
                    def is_comparison_col(col):
                        col_clean = col.replace(' ', '').lower()
                        # Only consider it a comparison column if it has total_price AND contains brackets or underscores
                        # This excludes plain 'total_price' columns which are single-offer
                        return (col_clean.startswith('total_price[') and ']' in col_clean) or \
                               (col_clean.startswith('total_price_') and '_' in col_clean[12:])  # After 'total_price_'
                    
                    if any(is_comparison_col(col) for col in df.columns):
                        # Comparison dataset: multiple offers
                        offer_columns = {}
                        for col in df.columns:
                            col_clean = col.replace(' ', '').lower()
                            if col_clean.startswith('total_price[') and ']' in col:
                                offer_key = col[col.find('[')+1:col.find(']')]
                                offer_columns[offer_key] = col
                            elif col_clean.startswith('total_price_'):
                                offer_key = col.split('_', 1)[1]
                                offer_columns[offer_key] = col
                        
                        offers_info = file_data.get('offers', {})
                        print(f"[DEBUG] Offers in file {file_key}: {list(offers_info.keys())}")
                        
                        for offer_key, price_col in offer_columns.items():
                            offer_info = offers_info.get(offer_key, {})
                            supplier = offer_info.get('supplier_name', offer_key)
                            project_name = offer_info.get('project_name', 'Unknown')
                            date = offer_info.get('date', 'Unknown')
                            total_price = self._calculate_total_price(file_mapping, offer_name=offer_key)
                            print(f"[DEBUG] Summary row: offer_key={offer_key}, supplier={supplier}, project={project_name}, date={date}, total_price={total_price}")
                            summary_data.append([supplier, project_name, date, total_price])
                    else:
                        # Single-offer dataset (legacy)
                        offer_info = file_data.get('offer_info', {})
                        print(f"[DEBUG] Single-offer dataset - file_data offer_info: {offer_info}")
                        
                        # Enhanced fallback logic for single-offer datasets
                        supplier = offer_info.get('supplier_name', 'Unknown')
                        project_name = offer_info.get('project_name', 'Unknown')
                        date = offer_info.get('date', 'Unknown')
                        
                        # If offer_info is empty or supplier is Unknown, try to get from current_offer_info
                        if not offer_info or supplier == 'Unknown':
                            current_offer_info = getattr(self, 'current_offer_info', {})
                            print(f"[DEBUG] Trying current_offer_info fallback: {current_offer_info}")
                            if current_offer_info:
                                supplier = current_offer_info.get('supplier_name', supplier)
                                project_name = current_offer_info.get('project_name', project_name)
                                date = current_offer_info.get('date', date)
                        
                        # If still Unknown, try to get from current_offer_name
                        if supplier == 'Unknown':
                            current_offer_name = getattr(self, 'current_offer_name', None)
                            print(f"[DEBUG] Trying current_offer_name fallback: {current_offer_name}")
                            if current_offer_name:
                                supplier = current_offer_name
                        
                        total_price = self._calculate_total_price(file_mapping, offer_name=supplier)
                        print(f"[DEBUG] Summary row (single): supplier={supplier}, project={project_name}, date={date}, total_price={total_price}")
                        summary_data.append([supplier, project_name, date, total_price])
        
        summary_data.sort(key=lambda x: float(x[3]) if x[3] != 'Unknown' else 0)
        return summary_data
        
        summary_data.sort(key=lambda x: float(x[3]) if x[3] != 'Unknown' else 0)
        return summary_data

    def _calculate_total_price(self, file_mapping, offer_name=None):
        """Calculate total price from file mapping, handling dynamic offer columns"""
        try:
            total_price = 0.0
            if hasattr(file_mapping, 'categorized_dataframe') and file_mapping.categorized_dataframe is not None:
                df = file_mapping.categorized_dataframe
            else:
                df = self._build_final_grid_dataframe(file_mapping)
            # Find total price column (dynamic)
            price_col = None
            if offer_name:
                # Try dynamic column name first
                for col in df.columns:
                    if col.lower().startswith('total_price') and offer_name in col:
                        price_col = col
                        break
            if not price_col:
                # Fallback to static column names
                for col in ['total_price', 'Total_price']:
                    if col in df.columns:
                        price_col = col
                        break
            if price_col and price_col in df.columns:
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
                df[price_col] = df[price_col].apply(parse_number)
                total_price = df[price_col].sum()
            return total_price
        except Exception as e:
            print(f"Error calculating total price: {e}")
            return 0.0

    def _copy_grid_to_clipboard(self, tree, grid_name):
        """Copy grid contents to clipboard in Excel-friendly format (point as thousands, comma as decimal)"""
        try:
            # Collect headers
            headers = [tree.heading(col)['text'] for col in tree['columns']]
            # Collect data rows
            data_rows = []
            for item in tree.get_children():
                values = tree.item(item)['values']
                data_rows.append(values)
            # Format as tab-separated values with Excel-compatible number formatting
            clipboard_text = '\t'.join(headers) + '\n'
            for row in data_rows:
                formatted_row = []
                for i, val in enumerate(row):
                    # Check if this is a numeric column (Total Price or similar)
                    if i < len(headers) and any(keyword in headers[i].lower() for keyword in ['price', 'total', 'amount', 'value']):
                        formatted_row.append(format_number_eu(val))
                    else:
                        formatted_row.append(str(val))
                clipboard_text += '\t'.join(formatted_row) + '\n'
            # Copy to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(clipboard_text)
            # Show confirmation
            messagebox.showinfo("Copied", f"{grid_name} data copied to clipboard")
        except Exception as e:
            print(f"Error copying to clipboard: {e}")
            messagebox.showerror("Error", f"Failed to copy {grid_name} data to clipboard")

    def _create_new_summary_grid(self, parent_frame, tab):
        """Create the new summary grid with Supplier, Project Name, Date, Total Price"""
        try:
            print(f"[DEBUG] _create_new_summary_grid called with tab type: {type(tab)}")
            print(f"[DEBUG] Tab object: {tab}")
            
            # Clear existing content
            for widget in parent_frame.winfo_children():
                widget.destroy()
            # Collect summary data from the current tab (not global files)
            summary_data = self._collect_summary_data(tab)
            if not summary_data:
                # No data to show
                no_data_label = ttk.Label(parent_frame, text="No BOQ data available for summary", 
                                        font=("TkDefaultFont", 10), foreground="gray")
                no_data_label.grid(row=0, column=0, pady=20)
                return
            # Create frame for the grid and copy button
            grid_frame = ttk.Frame(parent_frame)
            grid_frame.grid(row=0, column=0, sticky=tk.EW, pady=(5, 0))
            grid_frame.grid_columnconfigure(0, weight=1)
            # Create summary columns: Supplier, Project Name, Date, Total Price
            summary_columns = ['Supplier', 'Project Name', 'Date', 'Total Price']
            # Calculate height based on number of BOQs
            tree_height = max(2, len(summary_data))
            summary_tree = ttk.Treeview(grid_frame, columns=summary_columns, show='headings', height=tree_height)
            # Configure columns with appropriate widths
            column_widths = {
                'Supplier': 150,
                'Project Name': 200,
                'Date': 120,
                'Total Price': 150
            }
            for col in summary_columns:
                summary_tree.heading(col, text=col)
                summary_tree.column(col, width=column_widths[col], minwidth=100)
            # Insert data rows
            for row_data in summary_data:
                supplier, project_name, date, total_price = row_data
                # Format total price for display (point as thousands, comma as decimal)
                if total_price != 'Unknown' and total_price != 0.0:
                    display_price = format_number_eu(total_price)
                else:
                    display_price = 'Unknown'
                display_values = [supplier, project_name, date, display_price]
                summary_tree.insert('', 'end', values=display_values, tags=('boq_summary',))
            # Add horizontal scrollbar
            hsb_summary = ttk.Scrollbar(grid_frame, orient=tk.HORIZONTAL, command=summary_tree.xview)
            summary_tree.configure(xscrollcommand=hsb_summary.set)
            summary_tree.grid(row=0, column=0, sticky=tk.EW)
            hsb_summary.grid(row=1, column=0, sticky=tk.EW)
            # Create copy button frame
            copy_frame = ttk.Frame(parent_frame)
            copy_frame.grid(row=1, column=0, sticky=tk.E, pady=(5, 0))
            # Add copy button with icon
            copy_button = ttk.Button(copy_frame, text="📋 Copy to Clipboard", 
                                   command=lambda: self._copy_grid_to_clipboard(summary_tree, "BOQ Summary"))
            copy_button.pack(side=tk.RIGHT)
            # Store reference to the tree for potential future use
            try:
                tab.new_summary_tree = summary_tree
            except AttributeError:
                # If tab doesn't support attribute assignment, just continue
                print(f"[DEBUG] Tab doesn't support attribute assignment, continuing without storing tree reference")
        except Exception as e:
            print(f"Error creating new summary grid: {e}")
            import traceback
            traceback.print_exc()

    def _refresh_new_summary_grid(self, tab):
        """Refresh the new summary grid when new BOQs are loaded"""
        try:
            # Find the new summary frame in the tab
            for widget in tab.winfo_children():
                if hasattr(widget, 'winfo_children'):
                    for child in widget.winfo_children():
                        if hasattr(child, 'winfo_children'):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, ttk.LabelFrame) and grandchild.cget('text') == "BOQ Summary Overview":
                                    # Found the summary frame, refresh it
                                    self._create_new_summary_grid(grandchild, tab)
                                    return
        except Exception as e:
            print(f"Error refreshing new summary grid: {e}")
            import traceback
            traceback.print_exc()

    def _refresh_summary_grid_centralized(self):
        """Centralized method to refresh the summary grid with comprehensive logging"""
        try:
            print(f"[DEBUG] _refresh_summary_grid_centralized called")
            print(f"[DEBUG] Current files count: {len(self.controller.current_files)}")
            
            # Check if global summary frame exists
            if not hasattr(self, 'global_summary_frame') or self.global_summary_frame is None:
                print(f"[DEBUG] Global summary frame does not exist yet")
                return
            
            # Get current tab widget (not just the ID)
            current_tab_id = self.notebook.select()
            if not current_tab_id:
                print(f"[DEBUG] No current tab selected")
                return
            
            current_tab = self.notebook.nametowidget(current_tab_id)
            print(f"[DEBUG] Refreshing summary grid for tab: {current_tab}")
            
            # Collect summary data from the current tab
            summary_data = self._collect_summary_data(current_tab)
            print(f"[DEBUG] Collected {len(summary_data)} summary data entries")
            
            # Refresh the summary grid
            self._create_new_summary_grid(self.global_summary_frame, current_tab)
            print(f"[DEBUG] Summary grid refresh completed successfully")
            
        except Exception as e:
            print(f"[ERROR] Failed to refresh summary grid: {e}")
            import traceback
            traceback.print_exc()

    def _clear_all_files(self):
        """Clear all processed files and refresh summary grid"""
        try:
            print(f"[DEBUG] _clear_all_files called")
            # Clear all processed files
            self.controller.current_files.clear()
            print(f"[DEBUG] Cleared {len(self.controller.current_files)} files")
            
            # Clear all tabs
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            print(f"[DEBUG] Cleared all tabs")
            
            # Refresh summary grid
            self._refresh_summary_grid_centralized()
            
            self._update_status("All files cleared")
            print(f"[DEBUG] Clear operation completed")
            
        except Exception as e:
            print(f"[ERROR] Failed to clear files: {e}")
            import traceback
            traceback.print_exc()
