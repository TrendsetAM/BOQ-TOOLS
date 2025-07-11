"""
Offer Information Dialog for BOQ Tools
Comprehensive dialog for collecting offer information with specific fields and default behaviors
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Optional, Any
import platform
from datetime import datetime


class OfferInfoDialog:
    def __init__(self, parent, is_first_boq: bool = True, previous_offer_info: Optional[Dict[str, Any]] = None):
        """
        Initialize the offer information dialog
        
        Args:
            parent: Parent window
            is_first_boq: True if this is the first BOQ being loaded, False for subsequent BOQs
            previous_offer_info: Previous offer information for default values in subsequent BOQs
        """
        self.parent = parent
        self.is_first_boq = is_first_boq
        self.previous_offer_info = previous_offer_info or {}
        self.result = None
        
        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Offer Information")
        self.dialog.geometry("500x400")
        self.dialog.minsize(450, 350)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog
        self._center_dialog()
        
        # Setup style
        self._setup_style()
        
        # Create widgets
        self._create_widgets()
        
        # Set focus
        self.dialog.focus_set()
        
        # Bind events
        self.dialog.bind('<Return>', lambda e: self._on_confirm())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _center_dialog(self):
        """Center the dialog on the parent window"""
        self.dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
    
    def _setup_style(self):
        """Setup dialog style"""
        style = ttk.Style(self.dialog)
        if platform.system() == 'Windows':
            style.theme_use('vista')
        elif platform.system() == 'Darwin':
            style.theme_use('aqua')
        else:
            style.theme_use('clam')
    
    def _create_widgets(self):
        """Create all dialog widgets"""
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_text = "First BOQ Information" if self.is_first_boq else "Additional BOQ Information"
        title_label = ttk.Label(main_frame, text=title_text, 
                               font=("TkDefaultFont", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        if self.is_first_boq:
            instructions = "Please provide the following information for your BOQ analysis:"
        else:
            instructions = "Please provide information for the additional BOQ to compare:"
        
        instruction_label = ttk.Label(main_frame, text=instructions, 
                                    wraplength=450, justify=tk.LEFT)
        instruction_label.pack(pady=(0, 20))
        
        # Form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        form_frame.grid_columnconfigure(1, weight=1)
        
        # Supplier Name (Required) - Always first
        row = 0
        ttk.Label(form_frame, text="Supplier Name:", font=("TkDefaultFont", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=(0, 10), padx=(0, 10))
        
        self.supplier_var = tk.StringVar()
        supplier_entry = ttk.Entry(form_frame, textvariable=self.supplier_var, font=("TkDefaultFont", 10))
        supplier_entry.grid(row=row, column=1, sticky=tk.EW, pady=(0, 10))
        
        # Required indicator
        ttk.Label(form_frame, text="*", foreground="red", font=("TkDefaultFont", 10, "bold")).grid(
            row=row, column=2, sticky=tk.W, pady=(0, 10), padx=(5, 0))
        
        # Set default for supplier (empty for subsequent BOQs)
        if self.is_first_boq:
            self.supplier_var.set("")
        else:
            self.supplier_var.set("")  # Always blank for subsequent BOQs
        
        # Project Name (Optional)
        row += 1
        ttk.Label(form_frame, text="Project Name:").grid(
            row=row, column=0, sticky=tk.W, pady=(0, 10), padx=(0, 10))
        
        self.project_name_var = tk.StringVar()
        project_name_entry = ttk.Entry(form_frame, textvariable=self.project_name_var)
        project_name_entry.grid(row=row, column=1, sticky=tk.EW, pady=(0, 10))
        
        # Set default for project name
        if self.is_first_boq:
            self.project_name_var.set("")  # Blank for first BOQ
        else:
            # Use previous value for subsequent BOQs
            self.project_name_var.set(self.previous_offer_info.get('project_name', ''))
        
        # Project Size (Optional)
        row += 1
        ttk.Label(form_frame, text="Project Size:").grid(
            row=row, column=0, sticky=tk.W, pady=(0, 10), padx=(0, 10))
        
        self.project_size_var = tk.StringVar()
        project_size_entry = ttk.Entry(form_frame, textvariable=self.project_size_var)
        project_size_entry.grid(row=row, column=1, sticky=tk.EW, pady=(0, 10))
        
        # Set default for project size
        if self.is_first_boq:
            self.project_size_var.set("")  # Blank for first BOQ
        else:
            # Use previous value for subsequent BOQs
            self.project_size_var.set(self.previous_offer_info.get('project_size', ''))
        
        # Date (Optional)
        row += 1
        ttk.Label(form_frame, text="Date:").grid(
            row=row, column=0, sticky=tk.W, pady=(0, 10), padx=(0, 10))
        
        self.date_var = tk.StringVar()
        date_entry = ttk.Entry(form_frame, textvariable=self.date_var)
        date_entry.grid(row=row, column=1, sticky=tk.EW, pady=(0, 10))
        
        # Set default date to current date (July 11, 2025) for both first and subsequent BOQs
        default_date = "2025-07-11"  # As specified in requirements
        self.date_var.set(default_date)
        
        # Optional fields note
        row += 1
        note_label = ttk.Label(form_frame, text="* Required field. Other fields are optional.", 
                              font=("TkDefaultFont", 9), foreground="gray")
        note_label.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=(20, 0))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(30, 0))
        
        # Buttons
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Confirm", command=self._on_confirm).pack(side=tk.RIGHT)
        
        # Set initial focus to supplier field
        supplier_entry.focus_set()
    
    def _on_confirm(self):
        """Handle confirm button click"""
        # Validate required fields
        supplier_name = self.supplier_var.get().strip()
        if not supplier_name:
            messagebox.showerror("Required Field", "Supplier Name is required and cannot be empty.", 
                               parent=self.dialog)
            return
        
        # Collect all information
        self.result = {
            'supplier_name': supplier_name,
            'project_name': self.project_name_var.get().strip(),
            'project_size': self.project_size_var.get().strip(),
            'date': self.date_var.get().strip(),
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Handle cancel button click"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict[str, str]]:
        """Show the dialog and return the result"""
        self.dialog.wait_window()
        return self.result


def show_offer_info_dialog(parent, is_first_boq: bool = True, 
                          previous_offer_info: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, str]]:
    """
    Show the offer information dialog
    
    Args:
        parent: Parent window
        is_first_boq: True if this is the first BOQ being loaded
        previous_offer_info: Previous offer information for defaults
    
    Returns:
        Dictionary with offer information or None if cancelled
    """
    dialog = OfferInfoDialog(parent, is_first_boq, previous_offer_info)
    return dialog.show() 