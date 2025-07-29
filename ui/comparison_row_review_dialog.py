#!/usr/bin/env python3
"""
Comparison Row Review Dialog
Shows a row review interface for comparison BoQ rows, allowing users to manually change validity
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
import pandas as pd

logger = logging.getLogger(__name__)


class ComparisonRowReviewDialog:
    def __init__(self, parent, file_mapping, row_results, offer_name="Comparison"):
        """
        Initialize the Comparison Row Review Dialog
        
        Args:
            parent: Parent window
            file_mapping: FileMapping object containing comparison data and column mappings
            row_results: List of row processing results from ComparisonProcessor
            offer_name: Name of the comparison offer
        """
        self.parent = parent
        self.file_mapping = file_mapping
        self.row_results = row_results
        self.offer_name = offer_name
        self.confirmed = False
        self.updated_results = None
        
        # Create the dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Row Review - {offer_name}")
        self.dialog.geometry("1200x700")
        
        # Enable maximize/minimize buttons and proper window controls
        self.dialog.resizable(True, True)
        self.dialog.minsize(800, 500)  # Set minimum size
        
        # Make dialog modal but allow window controls
        self.dialog.grab_set()
        
        # Ensure window controls are visible
        self.dialog.attributes('-toolwindow', False)
        
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
        
        # Get available columns from file mapping
        self.available_columns = self._get_available_columns()
        
        # Create treeview with same configuration as master row review
        columns = ['Row'] + self.available_columns + ['Status', 'Reason']
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12, selectmode="none")
        
        # Configure columns same as master
        for col in columns:
            self.tree.heading(col, text=col.capitalize() if col != '#' else '#')
            if col == "status":
                self.tree.column(col, width=80, anchor=tk.CENTER)
            else:
                self.tree.column(col, width=120 if col != "#" else 40, anchor=tk.W, minwidth=50, stretch=False)
        
        # Add scrollbars with proper configuration
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Use grid layout for treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Configure tags for row colors - same as master
        self.tree.tag_configure('validrow', background='#E8F5E9')  # light green
        self.tree.tag_configure('invalidrow', background='#FFEBEE')  # light red
        
        # Remove blue selection highlight - same as master
        style = ttk.Style(self.tree)
        style.map('Treeview', background=[('selected', '#FFEBEE')])  # Always red on select
        style.layout('Treeview.Item', [('Treeitem.padding', {'sticky': 'nswe', 'children': [('Treeitem.indicator', {'side': 'left', 'sticky': ''}), ('Treeitem.image', {'side': 'left', 'sticky': ''}), ('Treeitem.text', {'side': 'left', 'sticky': ''})]})])
        
        # Bind row click event
        self.tree.bind('<Button-1>', self._on_row_click)
        
        # Summary label
        self.summary_label = ttk.Label(main_frame, text="", font=("Arial", 10))
        self.summary_label.grid(row=3, column=0, pady=(10, 0), sticky=tk.W)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, pady=(10, 0), sticky=(tk.E, tk.W))
        button_frame.columnconfigure(1, weight=1)
        
        # Cancel button
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self._on_cancel)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Confirm button
        confirm_btn = ttk.Button(button_frame, text="Confirm", command=self._on_confirm)
        confirm_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
    def _get_available_columns(self):
        """Get available columns from file mapping"""
        available_columns = []
        
        # Define the desired column order (same as main dataset)
        desired_order = ['code', 'Source_Sheet', 'description', 'unit', 'quantity', 'unit_price', 'total_price', 'manhours', 'wage']
        
        # Check if file mapping has proper sheet structure with column mappings
        sheets = getattr(self.file_mapping, 'sheets', [])
        if sheets and any(hasattr(sheet, 'column_mappings') for sheet in sheets):
            # Use sheet structure with column mappings
            all_mapped_types = set()
            for sheet in sheets:
                if hasattr(sheet, 'column_mappings'):
                    for cm in sheet.column_mappings:
                        mapped_type = getattr(cm, 'mapped_type', None)
                        if mapped_type and mapped_type != 'ignore':
                            all_mapped_types.add(mapped_type)
            
            # Use desired order, but only include columns that are actually mapped
            for col in desired_order:
                if col in all_mapped_types:
                    available_columns.append(col)
            
            # Add any remaining mapped columns that weren't in the desired order
            for col in all_mapped_types:
                if col not in available_columns:
                    available_columns.append(col)
        else:
            # Fallback: use DataFrame columns directly
            dataframe = getattr(self.file_mapping, 'dataframe', None)
            if dataframe is not None:
                # Get columns from DataFrame
                df_columns = list(dataframe.columns)
                logger.info(f"DataFrame columns: {df_columns}")
                
                # Filter out unwanted columns (ignore_*, etc.)
                filtered_columns = []
                for col in df_columns:
                    if not (col.startswith('ignore') or 
                           col.startswith('_') or 
                           col in ['']):
                        filtered_columns.append(col)
                
                logger.info(f"Filtered columns: {filtered_columns}")
                
                # Map DataFrame column names to desired names (case-insensitive)
                column_mapping = {}
                for df_col in filtered_columns:
                    df_col_lower = df_col.lower()
                    for desired_col in desired_order:
                        if df_col_lower == desired_col.lower():
                            column_mapping[df_col] = desired_col
                            break
                    else:
                        # If no match found, keep original name
                        column_mapping[df_col] = df_col
                
                # Use desired order, but only include columns that exist in DataFrame
                for col in desired_order:
                    # Find matching DataFrame column (case-insensitive)
                    matching_df_col = None
                    for df_col in filtered_columns:
                        if df_col.lower() == col.lower():
                            matching_df_col = df_col
                            break
                    
                    if matching_df_col:
                        available_columns.append(col)
                
                # Add any remaining columns that weren't in the desired order
                for df_col in filtered_columns:
                    df_col_lower = df_col.lower()
                    if not any(df_col_lower == desired_col.lower() for desired_col in desired_order):
                        available_columns.append(df_col)
        
        logger.info(f"Available columns for comparison dialog: {available_columns}")
        return available_columns
        
    def _populate_treeview(self):
        """Populate the treeview with comparison data and row results"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Check if file mapping has proper sheet structure
        sheets = getattr(self.file_mapping, 'sheets', [])
        if sheets and any(hasattr(sheet, 'column_mappings') for sheet in sheets):
            # Use sheet structure approach
            self._populate_from_sheets()
        else:
            # Use DataFrame approach
            self._populate_from_dataframe()
        
        # Update summary
        self._update_summary()
    
    def _populate_from_sheets(self):
        """Populate treeview using sheet structure"""
        # Create a unified DataFrame from all sheets
        all_rows = []
        row_index_mapping = {}  # Map from unified index to (sheet_name, original_index)
        
        for sheet in getattr(self.file_mapping, 'sheets', []):
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
                    
                    # Create row dict with mapped columns
                    row_dict = {}
                    mapped_type_to_index = {}
                    
                    if hasattr(sheet, 'column_mappings'):
                        for cm in sheet.column_mappings:
                            mapped_type = getattr(cm, 'mapped_type', None)
                            if mapped_type and mapped_type != 'ignore':
                                mapped_type_to_index[mapped_type] = cm.column_index
                    
                    for col in self.available_columns:
                        idx = mapped_type_to_index.get(col)
                        val = row_data[idx] if idx is not None and idx < len(row_data) else ""
                        
                        # Apply number formatting for specific columns
                        if col in ['unit_price', 'total_price', 'wage']:
                            val = self._format_number(val, is_currency=True)
                        elif col in ['quantity', 'manhours']:
                            val = self._format_number(val, is_currency=False)
                            # Special formatting for manhours - only 2 decimals
                            if col == 'manhours' and val and val != '':
                                try:
                                    num_val = float(str(val).replace(',', '.'))
                                    val = f"{num_val:.2f}"
                                except:
                                    pass
                        
                        row_dict[col] = val
                    
                    all_rows.append(row_dict)
                    row_index_mapping[len(all_rows) - 1] = (sheet.sheet_name, rc.row_index)
        
        # Add rows to treeview
        for unified_idx, row_dict in enumerate(all_rows):
            # Find corresponding row result
            sheet_name, original_idx = row_index_mapping[unified_idx]
            row_result = None
            
            for result in self.row_results:
                if result['row_index'] == original_idx:
                    row_result = result
                    break
            
            if row_result is None:
                # Create default result if not found
                row_result = {
                    'row_index': original_idx,
                    'is_valid': True,
                    'reason': 'ROW_VALIDITY'
                }
            
            # Build row values
            row_values = [unified_idx + 1]  # Display 1-based row number
            for col in self.available_columns:
                row_values.append(row_dict.get(col, ''))
            
            # Add status and reason
            status = "Valid" if row_result['is_valid'] else "Invalid"
            reason = row_result.get('reason', 'ROW_VALIDITY')
            row_values.extend([status, reason])
            
            # Insert into treeview
            item = self.tree.insert('', 'end', values=row_values)
            
            # Set row color based on validity
            if row_result['is_valid']:
                self.tree.item(item, tags=('validrow',))
            else:
                self.tree.item(item, tags=('invalidrow',))
    
    def _populate_from_dataframe(self):
        """Populate treeview using DataFrame directly"""
        dataframe = getattr(self.file_mapping, 'dataframe', None)
        if dataframe is None:
            logger.error("No DataFrame found in file mapping")
            return
        
        # Create mapping from desired column names to actual DataFrame column names
        column_name_mapping = {}
        for desired_col in self.available_columns:
            for df_col in dataframe.columns:
                if df_col.lower() == desired_col.lower():
                    column_name_mapping[desired_col] = df_col
                    break
        
        logger.info(f"Column name mapping: {column_name_mapping}")
        
        # Add rows to treeview
        for idx, row in dataframe.iterrows():
            # Find corresponding row result
            row_result = None
            for result in self.row_results:
                if result['row_index'] == idx:
                    row_result = result
                    break
            
            if row_result is None:
                # Create default result if not found
                row_result = {
                    'row_index': idx,
                    'is_valid': True,
                    'reason': 'ROW_VALIDITY'
                }
            
            # Build row values
            row_values = [idx + 1]  # Display 1-based row number
            for col in self.available_columns:
                # Get the actual DataFrame column name
                df_col = column_name_mapping.get(col, col)
                val = row.get(df_col, '')
                
                # Apply number formatting for specific columns
                if col in ['unit_price', 'total_price', 'wage']:
                    val = self._format_number(val, is_currency=True)
                elif col in ['quantity', 'manhours']:
                    val = self._format_number(val, is_currency=False)
                    # Special formatting for manhours - only 2 decimals
                    if col == 'manhours' and val and val != '':
                        try:
                            num_val = float(str(val).replace(',', '.'))
                            val = f"{num_val:.2f}"
                        except:
                            pass
                
                row_values.append(val)
            
            # Add status and reason
            status = "Valid" if row_result['is_valid'] else "Invalid"
            reason = row_result.get('reason', 'ROW_VALIDITY')
            row_values.extend([status, reason])
            
            # Insert into treeview
            item = self.tree.insert('', 'end', values=row_values)
            
            # Set row color based on validity
            if row_result['is_valid']:
                self.tree.item(item, tags=('validrow',))
            else:
                self.tree.item(item, tags=('invalidrow',))
        
    def _format_number(self, value, is_currency=False):
        """Format number for display"""
        if not value or value == '':
            return ''
        
        try:
            # Convert to float
            if isinstance(value, str):
                value = float(value.replace(',', '.'))
            else:
                value = float(value)
            
            if is_currency:
                return f"€ {value:,.2f}".replace(',', ' ').replace('.', ',').replace(' ', '.')
            else:
                return f"{value:,.2f}".replace(',', ' ').replace('.', ',').replace(' ', '.')
        except:
            return str(value)
        
    def _on_row_click(self, event):
        """Handle row click to toggle validity - same as master row review"""
        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        
        # Get row values to extract the actual row index
        row_values = self.tree.item(row_id, 'values')
        if not row_values:
            return
        
        # Get the row index from the first column (which contains the display row number)
        try:
            idx = int(row_values[0]) - 1  # Convert to 0-based index
        except (ValueError, IndexError):
            return
        
        # Find the corresponding result
        row_result = None
        result_index = None
        for i, result in enumerate(self.row_results):
            if result['row_index'] == idx:
                row_result = result
                result_index = i
                break
        
        if row_result is None:
            # Create new result if not found
            row_result = {
                'row_index': idx,
                'is_valid': True,
                'reason': 'ROW_VALIDITY'
            }
            self.row_results.append(row_result)
            result_index = len(self.row_results) - 1
        
        # Toggle validity - same as master
        is_valid = self.row_results[result_index]['is_valid']
        new_valid = not is_valid
        self.row_results[result_index]['is_valid'] = new_valid
        
        # Update tag and status column - same as master
        tag = 'validrow' if new_valid else 'invalidrow'
        self.tree.item(row_id, tags=(tag,))
        vals = list(self.tree.item(row_id, 'values'))
        vals[-2] = "Valid" if new_valid else "Invalid"
        self.tree.item(row_id, values=vals)
        
        # Update summary
        self._update_summary()
    
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


def show_comparison_row_review(parent, file_mapping, row_results, offer_name="Comparison"):
    """
    Convenience function to show the comparison row review dialog
    
    Args:
        parent: Parent window
        file_mapping: FileMapping object containing comparison data and column mappings
        row_results: List of row processing results
        offer_name: Name of the comparison offer
        
    Returns:
        tuple: (confirmed, updated_results)
    """
    logger.info("show_comparison_row_review called")
    dialog = ComparisonRowReviewDialog(parent, file_mapping, row_results, offer_name)
    return dialog.show() 