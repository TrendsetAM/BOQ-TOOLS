#!/usr/bin/env python3
"""
Comparison Row Review Dialog
Shows a row review interface for comparison BoQ rows, allowing users to manually change validity
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging

logger = logging.getLogger(__name__)


class ComparisonRowReviewDialog:
    def __init__(self, parent, comparison_data, row_results, offer_name="Comparison"):
        """
        Initialize the Comparison Row Review Dialog
        
        Args:
            parent: Parent window
            comparison_data: DataFrame containing comparison data
            row_results: List of row processing results from ComparisonProcessor
            offer_name: Name of the comparison offer
        """
        self.parent = parent
        self.comparison_data = comparison_data
        self.row_results = row_results
        self.offer_name = offer_name
        self.confirmed = False
        self.updated_results = None
        
        # Create the dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Row Review - {offer_name}")
        self.dialog.geometry("1000x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        # Create the interface
        self._create_widgets()
        self._populate_treeview()
        
        # Bind events
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
    def _create_widgets(self):
        """Create the dialog widgets"""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title label
        title_label = ttk.Label(main_frame, text=f"Review {self.offer_name} Rows", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.W)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Click on a row to toggle its validity. Green = Valid, Red = Invalid")
        instructions.grid(row=1, column=0, pady=(0, 10), sticky=tk.W)
        
        # Create treeview frame
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Create treeview
        columns = ['Row', 'Description', 'Quantity', 'Unit_Price', 'Total_Price', 'Status', 'Reason']
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        self.tree.heading('Row', text='Row')
        self.tree.heading('Description', text='Description')
        self.tree.heading('Quantity', text='Quantity')
        self.tree.heading('Unit_Price', text='Unit Price')
        self.tree.heading('Total_Price', text='Total Price')
        self.tree.heading('Status', text='Status')
        self.tree.heading('Reason', text='Reason')
        
        # Set column widths
        self.tree.column('Row', width=50, minwidth=50)
        self.tree.column('Description', width=200, minwidth=150)
        self.tree.column('Quantity', width=80, minwidth=80)
        self.tree.column('Unit_Price', width=100, minwidth=100)
        self.tree.column('Total_Price', width=100, minwidth=100)
        self.tree.column('Status', width=80, minwidth=80)
        self.tree.column('Reason', width=150, minwidth=150)
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Bind click event
        self.tree.bind('<Button-1>', self._on_row_click)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=(10, 0), sticky=(tk.E, tk.W))
        button_frame.columnconfigure(1, weight=1)
        
        # Buttons
        self.cancel_btn = ttk.Button(button_frame, text="Cancel", command=self._on_cancel)
        self.cancel_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.confirm_btn = ttk.Button(button_frame, text="Confirm", command=self._on_confirm)
        self.confirm_btn.grid(row=0, column=2)
        
        # Summary frame
        summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding="5")
        summary_frame.grid(row=4, column=0, pady=(10, 0), sticky=(tk.E, tk.W))
        
        self.summary_label = ttk.Label(summary_frame, text="")
        self.summary_label.grid(row=0, column=0, sticky=tk.W)
        
    def _populate_treeview(self):
        """Populate the treeview with comparison data and row results"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add rows
        for i, result in enumerate(self.row_results):
            row_idx = result['row_index']
            row_data = self.comparison_data.iloc[row_idx]
            
            # Get row values
            description = str(row_data.get('Description', ''))
            quantity = str(row_data.get('Quantity', ''))
            unit_price = str(row_data.get('Unit_Price', ''))
            total_price = str(row_data.get('Total_Price', ''))
            
            # Format status and reason
            status = "Valid" if result['is_valid'] else "Invalid"
            reason = result['reason']
            
            # Insert into treeview
            item = self.tree.insert('', 'end', values=(
                row_idx + 1,  # Display 1-based row number
                description,
                quantity,
                unit_price,
                total_price,
                status,
                reason
            ))
            
            # Set row color based on validity
            if result['is_valid']:
                self.tree.tag_configure('valid', background='lightgreen')
                self.tree.item(item, tags=('valid',))
            else:
                self.tree.tag_configure('invalid', background='lightcoral')
                self.tree.item(item, tags=('invalid',))
        
        # Update summary
        self._update_summary()
        
    def _on_row_click(self, event):
        """Handle row click to toggle validity"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            selection = self.tree.selection()
            if not selection:
                return  # No item selected
            item = selection[0]
            row_idx = int(self.tree.item(item, 'values')[0]) - 1  # Convert back to 0-based
            
            # Find the corresponding result
            for i, result in enumerate(self.row_results):
                if result['row_index'] == row_idx:
                    # Toggle validity
                    self.row_results[i]['is_valid'] = not self.row_results[i]['is_valid']
                    
                    # Update the treeview
                    values = list(self.tree.item(item, 'values'))
                    values[5] = "Valid" if self.row_results[i]['is_valid'] else "Invalid"
                    self.tree.item(item, values=values)
                    
                    # Update row color
                    if self.row_results[i]['is_valid']:
                        self.tree.item(item, tags=('valid',))
                    else:
                        self.tree.item(item, tags=('invalid',))
                    
                    # Update summary
                    self._update_summary()
                    break
    
    def _update_summary(self):
        """Update the summary label"""
        total_rows = len(self.row_results)
        valid_rows = sum(1 for r in self.row_results if r['is_valid'])
        invalid_rows = total_rows - valid_rows
        
        summary_text = f"Total: {total_rows} | Valid: {valid_rows} | Invalid: {invalid_rows}"
        self.summary_label.config(text=summary_text)
    
    def _on_confirm(self):
        """Handle confirm button click"""
        self.confirmed = True
        self.updated_results = self.row_results.copy()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Handle cancel button click"""
        self.confirmed = False
        self.updated_results = None
        self.dialog.destroy()
    
    def show(self):
        """Show the dialog and wait for result"""
        self.dialog.wait_window()
        return self.confirmed, self.updated_results


def show_comparison_row_review(parent, comparison_data, row_results, offer_name="Comparison"):
    """
    Convenience function to show the comparison row review dialog
    
    Args:
        parent: Parent window
        comparison_data: DataFrame containing comparison data
        row_results: List of row processing results
        offer_name: Name of the comparison offer
        
    Returns:
        tuple: (confirmed, updated_results)
    """
    dialog = ComparisonRowReviewDialog(parent, comparison_data, row_results, offer_name)
    return dialog.show() 