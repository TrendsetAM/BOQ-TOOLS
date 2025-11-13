"""
Preview Dialog for BOQ Tools
Shows processing results and allows manual overrides
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional, Callable
import platform

# Color coding for confidence and status
def confidence_color(score: float) -> str:
    if score >= 0.8:
        return '#4CAF50'  # Green
    elif score >= 0.6:
        return '#FFC107'  # Amber
    else:
        return '#F44336'  # Red

def status_color(status: str) -> str:
    colors = {
        'success': '#4CAF50',
        'partial': '#FFC107', 
        'failed': '#F44336',
        'needs_review': '#FF9800'
    }
    return colors.get(status, '#757575')

def tooltip(widget, text: str):
    """Simple tooltip for a widget"""
    def on_enter(event):
        widget._tip = tk.Toplevel(widget)
        widget._tip.wm_overrideredirect(True)
        x = event.x_root + 10
        y = event.y_root + 10
        widget._tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(widget._tip, text=text, background="#ffffe0", 
                        relief='solid', borderwidth=1, font=("TkDefaultFont", 9))
        label.pack()
    def on_leave(event):
        if hasattr(widget, '_tip'):
            widget._tip.destroy()
            widget._tip = None
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)


class PreviewDialog:
    def __init__(self, parent, file_mapping: Dict[str, Any], on_confirm: Optional[Callable] = None):
        """
        Initialize the preview dialog
        
        Args:
            parent: Parent window
            file_mapping: File mapping data from MappingGenerator
            on_confirm: Callback function when user confirms changes
        """
        self.parent = parent
        self.file_mapping = file_mapping
        self.on_confirm = on_confirm
        self.user_changes = {}
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Preview Processing Results")
        self.dialog.geometry("1000x700")
        self.dialog.minsize(800, 600)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self._center_dialog()
        
        # Setup styling
        self._setup_style()
        
        # Create widgets
        self._create_widgets()
        
        # Bind keyboard shortcuts
        self._bind_shortcuts()
        
        # Populate data
        self._populate_data()
        
        # Focus on dialog
        self.dialog.focus_set()

    def _center_dialog(self):
        """Center the dialog on the parent window"""
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def _setup_style(self):
        """Setup consistent styling"""
        style = ttk.Style(self.dialog)
        if platform.system() == 'Windows':
            style.theme_use('vista')
        elif platform.system() == 'Darwin':
            style.theme_use('aqua')
        else:
            style.theme_use('clam')

    def _create_widgets(self):
        """Create the main widgets"""
        # Main frame
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Processing Results Preview", 
                               font=("TkDefaultFont", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self._create_summary_tab()
        self._create_sheets_tab()
        self._create_mappings_tab()
        self._create_preview_tab()
        
        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process")
        status_label = ttk.Label(button_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        # Buttons
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Confirm & Process", command=self._on_confirm).pack(side=tk.RIGHT, padx=5)
        
        # Set default button
        self.dialog.bind('<Return>', lambda e: self._on_confirm())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

    def _create_summary_tab(self):
        """Create the summary tab"""
        summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(summary_frame, text="Summary")
        
        # File info
        file_frame = ttk.LabelFrame(summary_frame, text="File Information", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        metadata = self.file_mapping.get('metadata', {})
        ttk.Label(file_frame, text=f"Filename: {metadata.get('filename', 'Unknown')}").pack(anchor=tk.W)
        ttk.Label(file_frame, text=f"Total Sheets: {metadata.get('total_sheets', 0)}").pack(anchor=tk.W)
        ttk.Label(file_frame, text=f"Processing Date: {metadata.get('processing_date', 'Unknown')}").pack(anchor=tk.W)
        
        # Global confidence
        conf_frame = ttk.LabelFrame(summary_frame, text="Overall Confidence", padding=10)
        conf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        global_conf = self.file_mapping.get('global_confidence', 0.0)
        conf_label = ttk.Label(conf_frame, text=f"{global_conf:.1%}", 
                              background=confidence_color(global_conf), foreground="white", padding=10)
        conf_label.pack()
        tooltip(conf_label, "Overall confidence score for the entire file")
        
        # Processing summary
        summary_data = self.file_mapping.get('processing_summary', {})
        summary_frame_inner = ttk.LabelFrame(summary_frame, text="Processing Summary", padding=10)
        summary_frame_inner.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create summary tree
        columns = ("Metric", "Value")
        self.summary_tree = ttk.Treeview(summary_frame_inner, columns=columns, show="headings", height=8)
        for col in columns:
            self.summary_tree.heading(col, text=col)
            self.summary_tree.column(col, width=150)
        
        summary_vsb = ttk.Scrollbar(summary_frame_inner, orient=tk.VERTICAL, command=self.summary_tree.yview)
        summary_hsb = ttk.Scrollbar(summary_frame_inner, orient=tk.HORIZONTAL, command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=summary_vsb.set, xscrollcommand=summary_hsb.set)
        
        self.summary_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        summary_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        summary_hsb.pack(side=tk.BOTTOM, fill=tk.X)

    def _create_sheets_tab(self):
        """Create the sheets tab"""
        sheets_frame = ttk.Frame(self.notebook)
        self.notebook.add(sheets_frame, text="Sheets")
        
        # Sheets tree
        columns = ("Sheet", "Type", "Status", "Confidence", "Rows", "Include")
        self.sheets_tree = ttk.Treeview(sheets_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            self.sheets_tree.heading(col, text=col)
            if col == "Include":
                self.sheets_tree.column(col, width=80, anchor=tk.CENTER)
            elif col == "Confidence":
                self.sheets_tree.column(col, width=100, anchor=tk.CENTER)
            else:
                self.sheets_tree.column(col, width=120)
        
        # Scrollbars
        sheets_vsb = ttk.Scrollbar(sheets_frame, orient=tk.VERTICAL, command=self.sheets_tree.yview)
        sheets_hsb = ttk.Scrollbar(sheets_frame, orient=tk.HORIZONTAL, command=self.sheets_tree.xview)
        self.sheets_tree.configure(yscrollcommand=sheets_vsb.set, xscrollcommand=sheets_hsb.set)
        
        self.sheets_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        sheets_vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        sheets_hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=5)
        
        # Bind double-click for sheet details
        self.sheets_tree.bind('<Double-1>', self._on_sheet_double_click)

    def _create_mappings_tab(self):
        """Create the column mappings tab"""
        mappings_frame = ttk.Frame(self.notebook)
        self.notebook.add(mappings_frame, text="Column Mappings")
        
        # Sheet selector
        selector_frame = ttk.Frame(mappings_frame)
        selector_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(selector_frame, text="Sheet:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(selector_frame, textvariable=self.sheet_var, state="readonly", width=20)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', self._on_sheet_selected)
        
        # Mappings tree
        columns = ("Column", "Original Header", "Mapped Type", "Confidence", "Override")
        self.mappings_tree = ttk.Treeview(mappings_frame, columns=columns, show="headings", height=12)
        
        for col in columns:
            self.mappings_tree.heading(col, text=col)
            if col == "Confidence":
                self.mappings_tree.column(col, width=100, anchor=tk.CENTER)
            elif col == "Override":
                self.mappings_tree.column(col, width=120, anchor=tk.CENTER)
            else:
                self.mappings_tree.column(col, width=150)
        
        # Scrollbars
        mappings_vsb = ttk.Scrollbar(mappings_frame, orient=tk.VERTICAL, command=self.mappings_tree.yview)
        mappings_hsb = ttk.Scrollbar(mappings_frame, orient=tk.HORIZONTAL, command=self.mappings_tree.xview)
        self.mappings_tree.configure(yscrollcommand=mappings_vsb.set, xscrollcommand=mappings_hsb.set)
        
        self.mappings_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        mappings_vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        mappings_hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=5)

    def _create_preview_tab(self):
        """Create the data preview tab"""
        preview_frame = ttk.Frame(self.notebook)
        self.notebook.add(preview_frame, text="Data Preview")
        
        # Sheet selector for preview
        preview_selector_frame = ttk.Frame(preview_frame)
        preview_selector_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(preview_selector_frame, text="Preview Sheet:").pack(side=tk.LEFT)
        self.preview_sheet_var = tk.StringVar()
        self.preview_sheet_combo = ttk.Combobox(preview_selector_frame, textvariable=self.preview_sheet_var, 
                                               state="readonly", width=20)
        self.preview_sheet_combo.pack(side=tk.LEFT, padx=5)
        self.preview_sheet_combo.bind('<<ComboboxSelected>>', self._on_preview_sheet_selected)
        
        # Data preview tree
        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=15)
        
        # Scrollbars
        preview_v_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        preview_h_scroll = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_v_scroll.set, xscrollcommand=preview_h_scroll.set)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        preview_v_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        preview_h_scroll.pack(side=tk.BOTTOM, fill=tk.X, padx=5)

    def _bind_shortcuts(self):
        """Bind keyboard shortcuts"""
        self.dialog.bind('<Control-s>', lambda e: self._on_confirm())
        self.dialog.bind('<Control-c>', lambda e: self._on_cancel())
        self.dialog.bind('<F5>', lambda e: self._refresh_data())

    def _populate_data(self):
        """Populate the dialog with data"""
        # Populate summary
        self._populate_summary()
        
        # Populate sheets
        self._populate_sheets()
        
        # Populate sheet selectors
        self._populate_sheet_selectors()
        
        # Select first sheet by default
        if self.sheet_var.get():
            self._on_sheet_selected(None)

    def _populate_summary(self):
        """Populate summary tab"""
        summary_data = self.file_mapping.get('processing_summary', {})
        
        summary_items = [
            ("Total Rows Processed", str(summary_data.get('total_rows_processed', 0))),
            ("Total Columns Mapped", str(summary_data.get('total_columns_mapped', 0))),
            ("Successful Sheets", str(summary_data.get('successful_sheets', 0))),
            ("Partial Sheets", str(summary_data.get('partial_sheets', 0))),
            ("Failed Sheets", str(summary_data.get('failed_sheets', 0))),
            ("Sheets Needing Review", str(summary_data.get('sheets_needing_review', 0))),
            ("Average Confidence", f"{summary_data.get('average_confidence', 0.0):.1%}"),
            ("Total Validation Errors", str(summary_data.get('total_validation_errors', 0))),
            ("Total Validation Warnings", str(summary_data.get('total_validation_warnings', 0))),
        ]
        
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        
        for metric, value in summary_items:
            self.summary_tree.insert('', tk.END, values=(metric, value))

    def _populate_sheets(self):
        """Populate sheets tab"""
        sheets = self.file_mapping.get('sheets', [])
        
        for item in self.sheets_tree.get_children():
            self.sheets_tree.delete(item)
        
        for sheet in sheets:
            sheet_name = sheet.get('sheet_name', 'Unknown')
            sheet_type = sheet.get('classification_type', 'Unknown')
            status = sheet.get('processing_status', 'unknown')
            confidence = sheet.get('overall_confidence', 0.0)
            row_count = sheet.get('row_count', 0)
            
            # Create checkbox for inclusion
            include_var = tk.BooleanVar(value=True)
            
            item = self.sheets_tree.insert('', tk.END, values=(
                sheet_name, sheet_type, status, f"{confidence:.1%}", row_count, "✓"
            ))
            
            # Store the checkbox variable
            self.sheets_tree.set(item, "Include", "✓")

    def _populate_sheet_selectors(self):
        """Populate sheet selectors"""
        sheets = self.file_mapping.get('sheets', [])
        sheet_names = [sheet.get('sheet_name', 'Unknown') for sheet in sheets]
        
        self.sheet_combo['values'] = sheet_names
        self.preview_sheet_combo['values'] = sheet_names
        
        if sheet_names:
            self.sheet_combo.current(0)
            self.preview_sheet_combo.current(0)

    def _on_sheet_selected(self, event):
        """Handle sheet selection in mappings tab"""
        sheet_name = self.sheet_var.get()
        if not sheet_name:
            return
        
        # Find the selected sheet
        sheets = self.file_mapping.get('sheets', [])
        selected_sheet = None
        for sheet in sheets:
            if sheet.get('sheet_name') == sheet_name:
                selected_sheet = sheet
                break
        
        if not selected_sheet:
            return
        
        # Clear existing mappings
        for item in self.mappings_tree.get_children():
            self.mappings_tree.delete(item)
        
        # Populate column mappings
        column_mappings = selected_sheet.get('column_mappings', [])
        for mapping in column_mappings:
            col_index = mapping.get('column_index', 0)
            original_header = mapping.get('original_header', '')
            mapped_type = mapping.get('mapped_type', '')
            confidence = mapping.get('confidence', 0.0)
            
            item = self.mappings_tree.insert('', tk.END, values=(
                f"Col {col_index}", original_header, mapped_type, f"{confidence:.1%}", "Edit"
            ))
            
            # Color code confidence
            if confidence < 0.6:
                self.mappings_tree.tag_configure('low_conf', background='#ffebee')
                self.mappings_tree.item(item, tags=('low_conf',))

    def _on_preview_sheet_selected(self, event):
        """Handle sheet selection in preview tab"""
        sheet_name = self.preview_sheet_var.get()
        if not sheet_name:
            return
        
        # Find the selected sheet data
        sheet_data = self.file_mapping.get('sheet_data', {}).get(sheet_name, [])
        
        # Clear existing preview
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        if not sheet_data:
            return
        
        # Set up columns
        if sheet_data:
            columns = [f"Col {i}" for i in range(len(sheet_data[0]))]
            self.preview_tree['columns'] = columns
            
            for col in columns:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100)
        
        # Populate data (first 20 rows)
        for i, row in enumerate(sheet_data[:20]):
            self.preview_tree.insert('', tk.END, values=row)

    def _on_sheet_double_click(self, event):
        """Handle double-click on sheet in sheets tab"""
        selection = self.sheets_tree.selection()
        if selection:
            item = selection[0]
            sheet_name = self.sheets_tree.item(item, 'values')[0]
            
            # Switch to mappings tab and select this sheet
            self.notebook.select(1)  # Mappings tab
            self.sheet_var.set(sheet_name)
            self._on_sheet_selected(None)

    def _refresh_data(self):
        """Refresh the displayed data"""
        self._populate_data()
        self.status_var.set("Data refreshed")

    def _on_confirm(self):
        """Handle confirm button click"""
        # Collect user changes
        changes = {
            'sheet_inclusions': {},
            'column_overrides': {},
            'user_notes': []
        }
        
        # Get sheet inclusions
        for item in self.sheets_tree.get_children():
            values = self.sheets_tree.item(item, 'values')
            sheet_name = values[0]
            included = values[5] == "✓"
            changes['sheet_inclusions'][sheet_name] = included
        
        # Store changes
        self.user_changes = changes
        
        # Call callback if provided
        if self.on_confirm:
            self.on_confirm(changes)
        
        self.dialog.destroy()

    def _on_cancel(self):
        """Handle cancel button click"""
        self.dialog.destroy()

    def show(self):
        """Show the dialog and return user changes"""
        self.dialog.wait_window()
        return self.user_changes


# Convenience function
def show_preview_dialog(parent, file_mapping: Dict[str, Any], on_confirm: Optional[Callable] = None) -> Dict[str, Any]:
    """
    Show the preview dialog
    
    Args:
        parent: Parent window
        file_mapping: File mapping data
        on_confirm: Optional callback function
        
    Returns:
        Dictionary of user changes
    """
    dialog = PreviewDialog(parent, file_mapping, on_confirm)
    return dialog.show() 