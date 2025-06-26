"""
Category Review Dialog for BOQ Tools
Allows users to review and modify categories before finalizing
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from typing import Dict, List, Any, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class CategoryReviewDialog:
    def __init__(self, parent, dataframe, on_save=None):
        """
        Initialize the category review dialog
        
        Args:
            parent: Parent window
            dataframe: DataFrame with categorization data
            on_save: Callback function when categories are saved
        """
        self.parent = parent
        self.dataframe = dataframe.copy()  # Work with a copy
        self.original_dataframe = dataframe
        self.on_save = on_save
        
        # Dialog state
        self.dialog = None
        self.tree = None
        self.category_var = tk.StringVar()
        self.filter_var = tk.StringVar()
        
        # Get unique categories
        self.categories = self._get_unique_categories()
        
        # Create and show dialog
        self._create_dialog()
    
    def _get_unique_categories(self) -> List[str]:
        """Get unique categories from the DataFrame"""
        if 'Category' in self.dataframe.columns:
            categories = self.dataframe['Category'].dropna().unique().tolist()
            return sorted([cat for cat in categories if cat and str(cat).strip()])
        return []
    
    def _create_dialog(self):
        """Create the main dialog window"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Category Review")
        self.dialog.geometry("1000x700")
        self.dialog.minsize(800, 600)
        
        # Make dialog modal
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog on parent
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"1000x700+{x}+{y}")
        
        # Create main content
        self._create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _create_widgets(self):
        """Create the dialog widgets"""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=tk.NSEW)
        
        # Configure grid weights
        self.dialog.grid_rowconfigure(0, weight=1)
        self.dialog.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Title and instructions
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 10))
        
        title_label = ttk.Label(title_frame, text="Review and Modify Categories", 
                               font=("TkDefaultFont", 14, "bold"))
        title_label.pack(side=tk.LEFT)
        
        instruction_label = ttk.Label(title_frame, 
                                     text="Double-click a row to edit its category", 
                                     font=("TkDefaultFont", 9))
        instruction_label.pack(side=tk.RIGHT)
        
        # Filter frame
        filter_frame = ttk.LabelFrame(main_frame, text="Filters", padding="5")
        filter_frame.grid(row=1, column=0, sticky=tk.EW, pady=(0, 10))
        
        # Category filter
        ttk.Label(filter_frame, text="Filter by Category:").grid(row=0, column=0, padx=(0, 5))
        category_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                     values=['All'] + self.categories, state='readonly')
        category_combo.grid(row=0, column=1, padx=(0, 10))
        category_combo.set('All')
        category_combo.bind('<<ComboboxSelected>>', self._on_filter_change)
        
        # Statistics
        stats_frame = ttk.Frame(filter_frame)
        stats_frame.grid(row=0, column=2, padx=(20, 0))
        
        self.stats_label = ttk.Label(stats_frame, text="")
        self.stats_label.pack()
        
        # Configure filter frame grid
        filter_frame.grid_columnconfigure(2, weight=1)
        
        # Treeview frame
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Create treeview
        self._create_treeview(tree_frame)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=tk.EW, pady=(10, 0))
        
        # Left side buttons
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=tk.LEFT)
        
        refresh_button = ttk.Button(left_buttons, text="Refresh", 
                                   command=self._refresh_data)
        refresh_button.pack(side=tk.LEFT, padx=(0, 5))
        
        export_button = ttk.Button(left_buttons, text="Export Changes", 
                                  command=self._export_changes)
        export_button.pack(side=tk.LEFT)
        
        # Right side buttons
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side=tk.RIGHT)
        
        cancel_button = ttk.Button(right_buttons, text="Cancel", 
                                  command=self._on_cancel)
        cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        save_button = ttk.Button(right_buttons, text="Save Changes", 
                                command=self._on_save)
        save_button.pack(side=tk.RIGHT)
        
        # Configure button frame
        button_frame.grid_columnconfigure(0, weight=1)
        
        # Initial load
        self._load_data()
        self._update_statistics()
    
    def _create_treeview(self, parent):
        """Create the treeview for displaying data"""
        # Create treeview with scrollbars
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=0, column=0, sticky=tk.NSEW)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        
        # Treeview
        columns = ('Description', 'Category', 'Source_Sheet', 'Row_Number')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        self.tree.heading('Description', text='Description')
        self.tree.heading('Category', text='Category')
        self.tree.heading('Source_Sheet', text='Source Sheet')
        self.tree.heading('Row_Number', text='Row #')
        
        self.tree.column('Description', width=400, minwidth=200)
        self.tree.column('Category', width=150, minwidth=100)
        self.tree.column('Source_Sheet', width=120, minwidth=80)
        self.tree.column('Row_Number', width=80, minwidth=60)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        h_scrollbar.grid(row=1, column=0, sticky=tk.EW)
        
        # Bind events
        self.tree.bind('<Double-1>', self._on_double_click)
        self.tree.bind('<ButtonRelease-1>', self._on_selection_change)
    
    def _load_data(self):
        """Load data into the treeview"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get filter value
        filter_category = self.filter_var.get()
        
        # Load data
        for index, row in self.dataframe.iterrows():
            description = str(row.get('Description', ''))[:100]  # Truncate long descriptions
            category = str(row.get('Category', ''))
            source_sheet = str(row.get('Source_Sheet', ''))
            row_number = index + 1
            
            # Apply filter
            if filter_category != 'All' and category != filter_category:
                continue
            
            # Insert into treeview
            item = self.tree.insert('', 'end', values=(description, category, source_sheet, row_number))
            
            # Store the original index for reference
            self.tree.set(item, 'index', index)
    
    def _update_statistics(self):
        """Update statistics display"""
        total_rows = len(self.dataframe)
        categorized_rows = len(self.dataframe[self.dataframe['Category'].notna() & 
                                             (self.dataframe['Category'] != '')])
        uncategorized_rows = total_rows - categorized_rows
        
        if 'Category' in self.dataframe.columns:
            unique_categories = len(self.dataframe['Category'].dropna().unique())
        else:
            unique_categories = 0
        
        stats_text = f"Total: {total_rows} | Categorized: {categorized_rows} | Uncategorized: {uncategorized_rows} | Categories: {unique_categories}"
        self.stats_label.config(text=stats_text)
    
    def _on_filter_change(self, event=None):
        """Handle filter change"""
        self._load_data()
        self._update_statistics()
    
    def _on_double_click(self, event):
        """Handle double-click on treeview item"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        if not values:
            return
        
        # Get current category
        current_category = values[1]
        
        # Show category selection dialog
        new_category = self._show_category_dialog(current_category)
        if new_category is not None:
            # Update the treeview
            self.tree.set(item, 'Category', new_category)
            
            # Update the DataFrame
            index = self.tree.set(item, 'index')
            if index is not None:
                self.dataframe.at[int(index), 'Category'] = new_category
            
            # Update statistics
            self._update_statistics()
    
    def _show_category_dialog(self, current_category):
        """Show dialog for selecting a new category"""
        dialog = tk.Toplevel(self.dialog)
        dialog.title("Select Category")
        dialog.geometry("400x300")
        dialog.transient(self.dialog)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label
        ttk.Label(main_frame, text="Select a new category:", 
                 font=("TkDefaultFont", 10, "bold")).pack(pady=(0, 10))
        
        # Category selection
        category_var = tk.StringVar(value=current_category)
        
        # Combobox with existing categories
        ttk.Label(main_frame, text="Existing categories:").pack(anchor=tk.W)
        category_combo = ttk.Combobox(main_frame, textvariable=category_var, 
                                     values=self.categories, state='readonly')
        category_combo.pack(fill=tk.X, pady=(5, 15))
        
        # New category entry
        ttk.Label(main_frame, text="Or enter a new category:").pack(anchor=tk.W)
        new_category_entry = ttk.Entry(main_frame, textvariable=category_var)
        new_category_entry.pack(fill=tk.X, pady=(5, 15))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        result = [None]  # Use list to store result
        
        def on_ok():
            result[0] = category_var.get().strip()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)
        
        # Focus on entry
        new_category_entry.focus_set()
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result[0]
    
    def _on_selection_change(self, event):
        """Handle selection change"""
        # Could be used to show details about selected row
        pass
    
    def _refresh_data(self):
        """Refresh the data display"""
        self._load_data()
        self._update_statistics()
    
    def _export_changes(self):
        """Export the current changes"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            title="Export Category Changes",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.dataframe.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    self.dataframe.to_excel(file_path, index=False)
                else:
                    self.dataframe.to_csv(file_path, index=False)
                
                messagebox.showinfo("Success", f"Changes exported to: {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export changes: {str(e)}")
    
    def _on_save(self):
        """Handle save button click"""
        # Check if there are changes
        if self.dataframe.equals(self.original_dataframe):
            messagebox.showinfo("Info", "No changes to save")
            return
        
        # Confirm save
        if messagebox.askyesno("Save Changes", 
                              "Are you sure you want to save the category changes?"):
            if self.on_save:
                self.on_save(self.dataframe)
            messagebox.showinfo("Success", "Category changes saved successfully!")
            self.dialog.destroy()
    
    def _on_cancel(self):
        """Handle cancel button click"""
        # Check if there are unsaved changes
        if not self.dataframe.equals(self.original_dataframe):
            if messagebox.askyesno("Unsaved Changes", 
                                  "You have unsaved changes. Are you sure you want to cancel?"):
                self.dialog.destroy()
        else:
            self.dialog.destroy()
    
    def _on_close(self):
        """Handle window close"""
        self._on_cancel()


def show_category_review_dialog(parent, dataframe, on_save=None):
    """
    Show the category review dialog
    
    Args:
        parent: Parent window
        dataframe: DataFrame with categorization data
        on_save: Callback function when categories are saved
    
    Returns:
        CategoryReviewDialog instance
    """
    dialog = CategoryReviewDialog(parent, dataframe, on_save)
    return dialog 