"""
Settings Dialog for BOQ Tools
Comprehensive settings management with validation and organization support
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Any, Optional, Callable
import json
import os
import platform
from pathlib import Path
import logging

# Default settings
DEFAULT_SETTINGS = {
    "user_preferences": {
        "default_export_format": "json",
        "default_export_location": str(Path.home() / "Documents" / "BOQ_Exports"),
        "processing_thresholds": {
            "confidence_minimum": 0.6,
            "validation_error_limit": 10,
            "missing_data_threshold": 0.3,
            "ambiguous_mapping_threshold": 0.7
        },
        "performance": {
            "max_memory_mb": 512,
            "chunk_size_rows": 1000,
            "max_concurrent_sheets": 4,
            "enable_caching": True
        }
    },
    "organization": {
        "company_name": "",
        "custom_column_names": {
            "item_number": ["Item No.", "Item Number", "Item #", "No."],
            "description": ["Description", "Specification", "Details", "Item Description"],
            "unit": ["Unit", "UOM", "Unit of Measure"],
            "quantity": ["Quantity", "Qty", "Amount", "Count"],
            "unit_price": ["Unit Price", "Rate", "Price", "Cost"],
            "total_price": ["Total Price", "Amount", "Total", "Sum"]
        },
        "boq_patterns": {
            "hierarchical_numbering": True,
            "section_breaks": ["Subtotal", "Total", "Summary"],
            "currency_symbols": ["$", "€", "£", "¥"],
            "common_units": ["m²", "m³", "kg", "pcs", "hr", "day"]
        },
        "default_classifications": {
            "line_items_keywords": ["Item", "Description", "Quantity", "Price"],
            "summary_keywords": ["Total", "Subtotal", "Summary", "Grand Total"],
            "reference_keywords": ["Notes", "Conditions", "Terms", "Contact"]
        }
    },
    "advanced": {
        "logging": {
            "level": "INFO",
            "file_location": str(Path.home() / "Documents" / "BOQ_Logs"),
            "max_file_size_mb": 10,
            "backup_count": 5,
            "console_output": True
        },
        "backup": {
            "auto_backup": True,
            "backup_interval_hours": 24,
            "backup_location": str(Path.home() / "Documents" / "BOQ_Backups"),
            "max_backups": 10
        },
        "import_export": {
            "auto_save_config": True,
            "config_file_location": str(Path.home() / "Documents" / "BOQ_Config"),
            "last_import_location": str(Path.home() / "Documents")
        }
    }
}

# Validation rules
VALIDATION_RULES = {
    "confidence_minimum": {"min": 0.0, "max": 1.0, "type": float},
    "validation_error_limit": {"min": 0, "max": 1000, "type": int},
    "missing_data_threshold": {"min": 0.0, "max": 1.0, "type": float},
    "ambiguous_mapping_threshold": {"min": 0.0, "max": 1.0, "type": float},
    "max_memory_mb": {"min": 64, "max": 8192, "type": int},
    "chunk_size_rows": {"min": 100, "max": 10000, "type": int},
    "max_concurrent_sheets": {"min": 1, "max": 16, "type": int},
    "max_file_size_mb": {"min": 1, "max": 100, "type": int},
    "backup_count": {"min": 1, "max": 50, "type": int},
    "backup_interval_hours": {"min": 1, "max": 168, "type": int},
    "max_backups": {"min": 1, "max": 100, "type": int}
}

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


class SettingsDialog:
    def __init__(self, parent, current_settings: Optional[Dict[str, Any]] = None, 
                 on_save: Optional[Callable] = None):
        """
        Initialize the settings dialog
        
        Args:
            parent: Parent window
            current_settings: Current settings dictionary
            on_save: Callback function when settings are saved
        """
        self.parent = parent
        self.current_settings = current_settings or DEFAULT_SETTINGS.copy()
        self.on_save = on_save
        self.original_settings = self.current_settings.copy()
        self.validation_errors = []
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Settings")
        self.dialog.geometry("800x600")
        self.dialog.minsize(700, 500)
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
        
        # Populate with current settings
        self._populate_settings()
        
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
        title_label = ttk.Label(main_frame, text="Settings", font=("TkDefaultFont", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self._create_user_preferences_tab()
        self._create_organization_tab()
        self._create_advanced_tab()
        
        # Bottom frame
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(bottom_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        # Buttons
        ttk.Button(bottom_frame, text="Reset to Defaults", command=self._reset_to_defaults).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Import", command=self._import_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Export", command=self._export_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Save", command=self._on_save).pack(side=tk.RIGHT, padx=5)

    def _create_user_preferences_tab(self):
        """Create the user preferences tab"""
        pref_frame = ttk.Frame(self.notebook)
        self.notebook.add(pref_frame, text="User Preferences")
        
        # Export settings
        export_frame = ttk.LabelFrame(pref_frame, text="Export Settings", padding=10)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Export format
        ttk.Label(export_frame, text="Default Export Format:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.export_format_var = tk.StringVar()
        export_format_combo = ttk.Combobox(export_frame, textvariable=self.export_format_var, 
                                          values=["json", "csv", "xlsx", "xml"], state="readonly", width=15)
        export_format_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(export_format_combo, "Default format for exported mappings")
        
        # Export location
        ttk.Label(export_frame, text="Default Export Location:").grid(row=1, column=0, sticky=tk.W, pady=2)
        export_location_frame = ttk.Frame(export_frame)
        export_location_frame.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        self.export_location_var = tk.StringVar()
        export_location_entry = ttk.Entry(export_location_frame, textvariable=self.export_location_var, width=40)
        export_location_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(export_location_frame, text="Browse", command=self._browse_export_location).pack(side=tk.RIGHT, padx=5)
        tooltip(export_location_entry, "Default folder for saving exported files")
        
        # Processing thresholds
        thresholds_frame = ttk.LabelFrame(pref_frame, text="Processing Thresholds", padding=10)
        thresholds_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Confidence minimum
        ttk.Label(thresholds_frame, text="Minimum Confidence:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.confidence_var = tk.DoubleVar()
        confidence_scale = ttk.Scale(thresholds_frame, from_=0.0, to=1.0, variable=self.confidence_var, 
                                   orient=tk.HORIZONTAL, length=200)
        confidence_scale.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        self.confidence_label = ttk.Label(thresholds_frame, text="0.6")
        self.confidence_label.grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        confidence_scale.configure(command=self._update_confidence_label)
        tooltip(confidence_scale, "Minimum confidence score for automatic processing")
        
        # Validation error limit
        ttk.Label(thresholds_frame, text="Max Validation Errors:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.error_limit_var = tk.IntVar()
        error_limit_spin = ttk.Spinbox(thresholds_frame, from_=0, to=1000, textvariable=self.error_limit_var, width=10)
        error_limit_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(error_limit_spin, "Maximum number of validation errors before flagging for review")
        
        # Missing data threshold
        ttk.Label(thresholds_frame, text="Missing Data Threshold:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.missing_data_var = tk.DoubleVar()
        missing_data_scale = ttk.Scale(thresholds_frame, from_=0.0, to=1.0, variable=self.missing_data_var, 
                                     orient=tk.HORIZONTAL, length=200)
        missing_data_scale.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        self.missing_data_label = ttk.Label(thresholds_frame, text="0.3")
        missing_data_label.grid(row=2, column=2, sticky=tk.W, padx=5, pady=2)
        missing_data_scale.configure(command=lambda v: self.missing_data_label.configure(text=f"{float(v):.1f}"))
        tooltip(missing_data_scale, "Threshold for flagging rows with missing required data")
        
        # Performance settings
        performance_frame = ttk.LabelFrame(pref_frame, text="Performance Settings", padding=10)
        performance_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Memory limit
        ttk.Label(performance_frame, text="Max Memory (MB):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.memory_var = tk.IntVar()
        memory_spin = ttk.Spinbox(performance_frame, from_=64, to=8192, textvariable=self.memory_var, width=10)
        memory_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(memory_spin, "Maximum memory usage for processing large files")
        
        # Chunk size
        ttk.Label(performance_frame, text="Chunk Size (Rows):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.chunk_size_var = tk.IntVar()
        chunk_spin = ttk.Spinbox(performance_frame, from_=100, to=10000, textvariable=self.chunk_size_var, width=10)
        chunk_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(chunk_spin, "Number of rows to process at once for large files")
        
        # Enable caching
        self.caching_var = tk.BooleanVar()
        caching_check = ttk.Checkbutton(performance_frame, text="Enable Caching", variable=self.caching_var)
        caching_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        tooltip(caching_check, "Cache processing results for faster repeated operations")

    def _create_organization_tab(self):
        """Create the organization settings tab"""
        org_frame = ttk.Frame(self.notebook)
        self.notebook.add(org_frame, text="Organization")
        
        # Company information
        company_frame = ttk.LabelFrame(org_frame, text="Company Information", padding=10)
        company_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(company_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.company_name_var = tk.StringVar()
        company_entry = ttk.Entry(company_frame, textvariable=self.company_name_var, width=40)
        company_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(company_entry, "Your organization name for configuration identification")
        
        # Custom column names
        column_frame = ttk.LabelFrame(org_frame, text="Custom Column Names", padding=10)
        column_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create column name editor
        self._create_column_editor(column_frame)
        
        # BOQ patterns
        patterns_frame = ttk.LabelFrame(org_frame, text="BOQ Patterns", padding=10)
        patterns_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Hierarchical numbering
        self.hierarchical_var = tk.BooleanVar()
        hierarchical_check = ttk.Checkbutton(patterns_frame, text="Enable Hierarchical Numbering", 
                                           variable=self.hierarchical_var)
        hierarchical_check.pack(anchor=tk.W, pady=2)
        tooltip(hierarchical_check, "Detect and use hierarchical item numbering (1.1, 1.2, etc.)")
        
        # Currency symbols
        ttk.Label(patterns_frame, text="Currency Symbols:").pack(anchor=tk.W, pady=(10, 2))
        self.currency_var = tk.StringVar()
        currency_entry = ttk.Entry(patterns_frame, textvariable=self.currency_var, width=40)
        currency_entry.pack(anchor=tk.W, pady=2)
        tooltip(currency_entry, "Comma-separated list of currency symbols to recognize")

    def _create_column_editor(self, parent):
        """Create the custom column names editor"""
        # Column types
        column_types = ["item_number", "description", "unit", "quantity", "unit_price", "total_price"]
        
        # Create treeview for editing
        columns = ("Column Type", "Alternative Names")
        self.column_tree = ttk.Treeview(parent, columns=columns, show="headings", height=6)
        
        for col in columns:
            self.column_tree.heading(col, text=col)
            if col == "Column Type":
                self.column_tree.column(col, width=120)
            else:
                self.column_tree.column(col, width=300)
        
        # Scrollbar
        column_scroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.column_tree.yview)
        self.column_tree.configure(yscrollcommand=column_scroll.set)
        
        self.column_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        column_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(button_frame, text="Edit Names", command=self._edit_column_names).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", command=self._reset_column_names).pack(side=tk.LEFT, padx=5)
        
        # Bind double-click
        self.column_tree.bind('<Double-1>', lambda e: self._edit_column_names())

    def _create_advanced_tab(self):
        """Create the advanced settings tab"""
        adv_frame = ttk.Frame(self.notebook)
        self.notebook.add(adv_frame, text="Advanced")
        
        # Logging settings
        logging_frame = ttk.LabelFrame(adv_frame, text="Logging", padding=10)
        logging_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Log level
        ttk.Label(logging_frame, text="Log Level:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.log_level_var = tk.StringVar()
        log_level_combo = ttk.Combobox(logging_frame, textvariable=self.log_level_var, 
                                      values=["DEBUG", "INFO", "WARNING", "ERROR"], state="readonly", width=15)
        log_level_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        tooltip(log_level_combo, "Level of detail for log messages")
        
        # Log file location
        ttk.Label(logging_frame, text="Log File Location:").grid(row=1, column=0, sticky=tk.W, pady=2)
        log_location_frame = ttk.Frame(logging_frame)
        log_location_frame.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        self.log_location_var = tk.StringVar()
        log_location_entry = ttk.Entry(log_location_frame, textvariable=self.log_location_var, width=40)
        log_location_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(log_location_frame, text="Browse", command=self._browse_log_location).pack(side=tk.RIGHT, padx=5)
        tooltip(log_location_entry, "Directory for storing log files")
        
        # Console output
        self.console_output_var = tk.BooleanVar()
        console_check = ttk.Checkbutton(logging_frame, text="Show Logs in Console", variable=self.console_output_var)
        console_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        tooltip(console_check, "Display log messages in the application console")
        
        # Backup settings
        backup_frame = ttk.LabelFrame(adv_frame, text="Backup & Recovery", padding=10)
        backup_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Auto backup
        self.auto_backup_var = tk.BooleanVar()
        auto_backup_check = ttk.Checkbutton(backup_frame, text="Enable Auto Backup", variable=self.auto_backup_var)
        auto_backup_check.pack(anchor=tk.W, pady=2)
        tooltip(auto_backup_check, "Automatically backup configuration files")
        
        # Backup interval
        ttk.Label(backup_frame, text="Backup Interval (hours):").pack(anchor=tk.W, pady=(10, 2))
        self.backup_interval_var = tk.IntVar()
        backup_interval_spin = ttk.Spinbox(backup_frame, from_=1, to=168, textvariable=self.backup_interval_var, width=10)
        backup_interval_spin.pack(anchor=tk.W, pady=2)
        tooltip(backup_interval_spin, "How often to create automatic backups")
        
        # Backup location
        ttk.Label(backup_frame, text="Backup Location:").pack(anchor=tk.W, pady=(10, 2))
        backup_location_frame = ttk.Frame(backup_frame)
        backup_location_frame.pack(fill=tk.X, pady=2)
        
        self.backup_location_var = tk.StringVar()
        backup_location_entry = ttk.Entry(backup_location_frame, textvariable=self.backup_location_var, width=40)
        backup_location_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(backup_location_frame, text="Browse", command=self._browse_backup_location).pack(side=tk.RIGHT, padx=5)
        tooltip(backup_location_entry, "Directory for storing backup files")

    def _bind_shortcuts(self):
        """Bind keyboard shortcuts"""
        self.dialog.bind('<Control-s>', lambda e: self._on_save())
        self.dialog.bind('<Control-r>', lambda e: self._reset_to_defaults())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

    def _populate_settings(self):
        """Populate the dialog with current settings"""
        # User preferences
        user_prefs = self.current_settings.get("user_preferences", {})
        self.export_format_var.set(user_prefs.get("default_export_format", "json"))
        self.export_location_var.set(user_prefs.get("default_export_location", ""))
        
        thresholds = user_prefs.get("processing_thresholds", {})
        self.confidence_var.set(thresholds.get("confidence_minimum", 0.6))
        self.error_limit_var.set(thresholds.get("validation_error_limit", 10))
        self.missing_data_var.set(thresholds.get("missing_data_threshold", 0.3))
        
        performance = user_prefs.get("performance", {})
        self.memory_var.set(performance.get("max_memory_mb", 512))
        self.chunk_size_var.set(performance.get("chunk_size_rows", 1000))
        self.caching_var.set(performance.get("enable_caching", True))
        
        # Organization
        org = self.current_settings.get("organization", {})
        self.company_name_var.set(org.get("company_name", ""))
        self.hierarchical_var.set(org.get("boq_patterns", {}).get("hierarchical_numbering", True))
        
        currency_symbols = org.get("boq_patterns", {}).get("currency_symbols", [])
        self.currency_var.set(", ".join(currency_symbols))
        
        # Populate column tree
        self._populate_column_tree()
        
        # Advanced
        advanced = self.current_settings.get("advanced", {})
        logging_settings = advanced.get("logging", {})
        self.log_level_var.set(logging_settings.get("level", "INFO"))
        self.log_location_var.set(logging_settings.get("file_location", ""))
        self.console_output_var.set(logging_settings.get("console_output", True))
        
        backup_settings = advanced.get("backup", {})
        self.auto_backup_var.set(backup_settings.get("auto_backup", True))
        self.backup_interval_var.set(backup_settings.get("backup_interval_hours", 24))
        self.backup_location_var.set(backup_settings.get("backup_location", ""))

    def _populate_column_tree(self):
        """Populate the column names tree"""
        for item in self.column_tree.get_children():
            self.column_tree.delete(item)
        
        custom_columns = self.current_settings.get("organization", {}).get("custom_column_names", {})
        for col_type, alternatives in custom_columns.items():
            alternatives_str = ", ".join(alternatives)
            self.column_tree.insert('', tk.END, values=(col_type.replace("_", " ").title(), alternatives_str))

    def _update_confidence_label(self, value):
        """Update confidence label"""
        self.confidence_label.configure(text=f"{float(value):.1f}")

    def _browse_export_location(self):
        """Browse for export location"""
        location = filedialog.askdirectory(title="Select Export Directory")
        if location:
            self.export_location_var.set(location)

    def _browse_log_location(self):
        """Browse for log location"""
        location = filedialog.askdirectory(title="Select Log Directory")
        if location:
            self.log_location_var.set(location)

    def _browse_backup_location(self):
        """Browse for backup location"""
        location = filedialog.askdirectory(title="Select Backup Directory")
        if location:
            self.backup_location_var.set(location)

    def _edit_column_names(self):
        """Edit column names in a dialog"""
        selection = self.column_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a column type to edit.")
            return
        
        item = selection[0]
        col_type = self.column_tree.item(item, 'values')[0]
        col_type_key = col_type.lower().replace(" ", "_")
        
        # Get current alternatives
        custom_columns = self.current_settings.get("organization", {}).get("custom_column_names", {})
        current_alternatives = custom_columns.get(col_type_key, [])
        
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.dialog)
        edit_dialog.title(f"Edit {col_type} Names")
        edit_dialog.geometry("400x300")
        edit_dialog.transient(self.dialog)
        edit_dialog.grab_set()
        
        ttk.Label(edit_dialog, text=f"Alternative names for {col_type}:").pack(pady=10)
        
        text_widget = tk.Text(edit_dialog, height=10, width=50)
        text_widget.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, ", ".join(current_alternatives))
        
        def save_changes():
            new_alternatives = [name.strip() for name in text_widget.get(1.0, tk.END).split(",") if name.strip()]
            custom_columns[col_type_key] = new_alternatives
            self._populate_column_tree()
            edit_dialog.destroy()
        
        ttk.Button(edit_dialog, text="Save", command=save_changes).pack(pady=10)

    def _reset_column_names(self):
        """Reset column names to defaults"""
        if messagebox.askyesno("Reset Column Names", "Reset all column names to defaults?"):
            self.current_settings["organization"]["custom_column_names"] = DEFAULT_SETTINGS["organization"]["custom_column_names"].copy()
            self._populate_column_tree()

    def _validate_settings(self) -> bool:
        """Validate all settings"""
        self.validation_errors = []
        
        # Validate numeric fields
        numeric_fields = [
            ("confidence_minimum", self.confidence_var.get()),
            ("validation_error_limit", self.error_limit_var.get()),
            ("missing_data_threshold", self.missing_data_var.get()),
            ("max_memory_mb", self.memory_var.get()),
            ("chunk_size_rows", self.chunk_size_var.get()),
            ("backup_interval_hours", self.backup_interval_var.get())
        ]
        
        for field_name, value in numeric_fields:
            if field_name in VALIDATION_RULES:
                rule = VALIDATION_RULES[field_name]
                if not isinstance(value, rule["type"]):
                    self.validation_errors.append(f"{field_name}: Invalid type")
                elif value < rule["min"] or value > rule["max"]:
                    self.validation_errors.append(f"{field_name}: Value must be between {rule['min']} and {rule['max']}")
        
        # Validate paths
        paths_to_check = [
            ("export_location", self.export_location_var.get()),
            ("log_location", self.log_location_var.get()),
            ("backup_location", self.backup_location_var.get())
        ]
        
        for path_name, path_value in paths_to_check:
            if path_value and not os.path.exists(path_value):
                try:
                    os.makedirs(path_value, exist_ok=True)
                except Exception:
                    self.validation_errors.append(f"{path_name}: Cannot create directory")
        
        return len(self.validation_errors) == 0

    def _collect_settings(self) -> Dict[str, Any]:
        """Collect settings from the dialog"""
        settings = {
            "user_preferences": {
                "default_export_format": self.export_format_var.get(),
                "default_export_location": self.export_location_var.get(),
                "processing_thresholds": {
                    "confidence_minimum": self.confidence_var.get(),
                    "validation_error_limit": self.error_limit_var.get(),
                    "missing_data_threshold": self.missing_data_var.get(),
                    "ambiguous_mapping_threshold": 0.7  # Default
                },
                "performance": {
                    "max_memory_mb": self.memory_var.get(),
                    "chunk_size_rows": self.chunk_size_var.get(),
                    "max_concurrent_sheets": 4,  # Default
                    "enable_caching": self.caching_var.get()
                }
            },
            "organization": {
                "company_name": self.company_name_var.get(),
                "custom_column_names": self.current_settings.get("organization", {}).get("custom_column_names", {}),
                "boq_patterns": {
                    "hierarchical_numbering": self.hierarchical_var.get(),
                    "section_breaks": ["Subtotal", "Total", "Summary"],
                    "currency_symbols": [s.strip() for s in self.currency_var.get().split(",") if s.strip()],
                    "common_units": ["m²", "m³", "kg", "pcs", "hr", "day"]
                },
                "default_classifications": DEFAULT_SETTINGS["organization"]["default_classifications"]
            },
            "advanced": {
                "logging": {
                    "level": self.log_level_var.get(),
                    "file_location": self.log_location_var.get(),
                    "max_file_size_mb": 10,
                    "backup_count": 5,
                    "console_output": self.console_output_var.get()
                },
                "backup": {
                    "auto_backup": self.auto_backup_var.get(),
                    "backup_interval_hours": self.backup_interval_var.get(),
                    "backup_location": self.backup_location_var.get(),
                    "max_backups": 10
                },
                "import_export": DEFAULT_SETTINGS["advanced"]["import_export"]
            }
        }
        
        return settings

    def _reset_to_defaults(self):
        """Reset all settings to defaults"""
        if messagebox.askyesno("Reset to Defaults", "Reset all settings to default values?"):
            self.current_settings = DEFAULT_SETTINGS.copy()
            self._populate_settings()
            self.status_var.set("Settings reset to defaults")

    def _import_settings(self):
        """Import settings from file"""
        filename = filedialog.askopenfilename(
            title="Import Settings",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported_settings = json.load(f)
                self.current_settings = imported_settings
                self._populate_settings()
                self.status_var.set(f"Settings imported from {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import settings: {e}")

    def _export_settings(self):
        """Export settings to file"""
        filename = filedialog.asksaveasfilename(
            title="Export Settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                settings = self._collect_settings()
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, indent=2, ensure_ascii=False)
                self.status_var.set(f"Settings exported to {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export settings: {e}")

    def _on_save(self):
        """Handle save button click"""
        if not self._validate_settings():
            error_msg = "Validation errors:\n" + "\n".join(self.validation_errors)
            messagebox.showerror("Validation Error", error_msg)
            return
        
        settings = self._collect_settings()
        self.current_settings = settings
        
        if self.on_save:
            self.on_save(settings)
        
        self.status_var.set("Settings saved successfully")
        self.dialog.destroy()

    def _on_cancel(self):
        """Handle cancel button click"""
        if self.current_settings != self.original_settings:
            if messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Discard them?"):
                self.dialog.destroy()
        else:
            self.dialog.destroy()

    def show(self) -> Dict[str, Any]:
        """Show the dialog and return the settings"""
        self.dialog.wait_window()
        return self.current_settings


# Convenience function
def show_settings_dialog(parent, current_settings: Optional[Dict[str, Any]] = None, 
                        on_save: Optional[Callable] = None) -> Dict[str, Any]:
    """
    Show the settings dialog
    
    Args:
        parent: Parent window
        current_settings: Current settings dictionary
        on_save: Optional callback function when settings are saved
        
    Returns:
        Dictionary of settings
    """
    dialog = SettingsDialog(parent, current_settings, on_save)
    return dialog.show() 