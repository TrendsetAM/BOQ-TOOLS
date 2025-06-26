"""
Categorization Dialog for BOQ Tools
Provides UI for the complete categorization workflow
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from pathlib import Path
from typing import Dict, Any, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class CategorizationDialog:
    def __init__(self, parent, controller, file_mapping, on_complete=None):
        """
        Initialize the categorization dialog
        
        Args:
            parent: Parent window
            controller: Main application controller
            file_mapping: File mapping object with processed data
            on_complete: Callback function when categorization is complete
        """
        self.parent = parent
        self.controller = controller
        self.file_mapping = file_mapping
        self.on_complete = on_complete
        
        # Dialog state
        self.dialog = None
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="Initializing categorization...")
        self.current_step = tk.StringVar(value="Step 1/6: Loading category dictionary")
        
        # Results
        self.categorization_result = None
        self.manual_excel_path = None
        self.final_dataframe = None
        
        # Create and show dialog
        self._create_dialog()
        self._start_categorization()
    
    def _create_dialog(self):
        """Create the main dialog window"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("BOQ Categorization Workflow")
        self.dialog.geometry("800x600")
        self.dialog.minsize(700, 500)
        
        # Make dialog modal
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog on parent
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)
        self.dialog.geometry(f"800x600+{x}+{y}")
        
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
        
        # Title
        title_label = ttk.Label(main_frame, text="BOQ Row Categorization", 
                               font=("TkDefaultFont", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=1, column=0, sticky=tk.NSEW, pady=(0, 10))
        
        # Current step
        step_label = ttk.Label(progress_frame, textvariable=self.current_step, 
                              font=("TkDefaultFont", 10, "bold"))
        step_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=400)
        self.progress_bar.grid(row=1, column=0, sticky=tk.EW, pady=(0, 5))
        
        # Status
        status_label = ttk.Label(progress_frame, textvariable=self.status_var, 
                                wraplength=500)
        status_label.grid(row=2, column=0, sticky=tk.W)
        
        # Configure progress frame grid
        progress_frame.grid_columnconfigure(0, weight=1)
        
        # Content area (changes based on current step)
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        self.content_frame.grid_columnconfigure(0, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=tk.EW, pady=(10, 0))
        
        self.cancel_button = ttk.Button(button_frame, text="Cancel", 
                                       command=self._on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.next_button = ttk.Button(button_frame, text="Next", 
                                     command=self._on_next, state=tk.DISABLED)
        self.next_button.pack(side=tk.RIGHT)
        
        self.back_button = ttk.Button(button_frame, text="Back", 
                                     command=self._on_back, state=tk.DISABLED)
        self.back_button.pack(side=tk.RIGHT, padx=(0, 5))
        
        # Configure button frame
        button_frame.grid_columnconfigure(0, weight=1)
    
    def _start_categorization(self):
        """Start the categorization process in a background thread"""
        def categorization_thread():
            try:
                from core.manual_categorizer import execute_row_categorization
                
                # Get the DataFrame from file mapping
                # This assumes the file mapping contains the processed DataFrame
                # You may need to adjust this based on your actual data structure
                mapped_df = self._get_dataframe_from_mapping()
                
                if mapped_df is None:
                    self._show_error("No data available for categorization")
                    return
                
                # Execute categorization with progress callback
                result = execute_row_categorization(
                    mapped_df=mapped_df,
                    progress_callback=self._update_progress
                )
                
                # Handle result
                if result['error']:
                    self._show_error(f"Categorization failed: {result['error']}")
                else:
                    self.categorization_result = result
                    self.final_dataframe = result['final_dataframe']
                    self._show_success()
                    
            except Exception as e:
                logger.error(f"Categorization error: {e}")
                self._show_error(f"Unexpected error: {str(e)}")
        
        # Start thread
        thread = threading.Thread(target=categorization_thread, daemon=True)
        thread.start()
    
    def _get_dataframe_from_mapping(self):
        """Extract DataFrame from file mapping"""
        # This is a placeholder - adjust based on your actual data structure
        try:
            # Assuming file_mapping has a method to get the processed DataFrame
            if hasattr(self.file_mapping, 'get_processed_dataframe'):
                return self.file_mapping.get_processed_dataframe()
            elif hasattr(self.file_mapping, 'dataframe'):
                return self.file_mapping.dataframe
            else:
                # Try to get from controller's current files
                return self.controller.get_current_dataframe()
        except Exception as e:
            logger.error(f"Error getting DataFrame: {e}")
            return None
    
    def _update_progress(self, percent, message):
        """Update progress bar and status"""
        self.progress_var.set(percent)
        self.status_var.set(message)
        
        # Update step based on progress
        if percent < 20:
            self.current_step.set("Step 1/6: Loading category dictionary")
        elif percent < 40:
            self.current_step.set("Step 2/6: Auto-categorizing rows")
        elif percent < 60:
            self.current_step.set("Step 3/6: Collecting unmatched descriptions")
        elif percent < 80:
            self.current_step.set("Step 4/6: Generating manual categorization file")
        elif percent < 95:
            self.current_step.set("Step 5/6: Processing manual categorizations")
        else:
            self.current_step.set("Step 6/6: Finalizing categorization")
    
    def _show_manual_categorization_step(self):
        """Show the manual categorization step"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Create manual categorization content
        manual_frame = ttk.LabelFrame(self.content_frame, text="Manual Categorization", padding="10")
        manual_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=10, pady=10)
        
        # Instructions
        instructions = """
        The system has generated a manual categorization Excel file with descriptions that couldn't be automatically categorized.
        
        To complete the categorization:
        1. Open the Excel file below
        2. Review each description and select an appropriate category from the dropdown
        3. Add any notes if needed
        4. Save the file
        5. Click 'Upload Completed File' to continue
        """
        
        instruction_label = ttk.Label(manual_frame, text=instructions, wraplength=500, justify=tk.LEFT)
        instruction_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Excel file path
        if self.categorization_result and 'summary' in self.categorization_result:
            excel_path = self.categorization_result['summary'].get('manual_excel_generated', '')
            if excel_path:
                path_label = ttk.Label(manual_frame, text=f"Excel file: {excel_path}")
                path_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(manual_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        open_button = ttk.Button(button_frame, text="Open Excel File", 
                                command=self._open_excel_file)
        open_button.pack(side=tk.LEFT, padx=(0, 5))
        
        upload_button = ttk.Button(button_frame, text="Upload Completed File", 
                                  command=self._upload_completed_file)
        upload_button.pack(side=tk.LEFT)
        
        # Configure grid weights
        manual_frame.grid_columnconfigure(0, weight=1)
    
    def _show_statistics_step(self):
        """Show categorization statistics"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        if not self.categorization_result:
            return
        
        # Create statistics content
        stats_frame = ttk.LabelFrame(self.content_frame, text="Categorization Statistics", padding="10")
        stats_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=10, pady=10)
        
        # Get statistics
        stats = self.categorization_result.get('all_stats', {})
        
        # Create statistics display
        row = 0
        
        # Auto-categorization stats
        auto_stats = stats.get('auto_stats', {})
        if auto_stats:
            ttk.Label(stats_frame, text="Auto-Categorization:", font=("TkDefaultFont", 10, "bold")).grid(
                row=row, column=0, sticky=tk.W, pady=(0, 5))
            row += 1
            
            ttk.Label(stats_frame, text=f"Total rows: {auto_stats.get('total_rows', 0)}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
            
            ttk.Label(stats_frame, text=f"Matched rows: {auto_stats.get('matched_rows', 0)}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
            
            ttk.Label(stats_frame, text=f"Match rate: {auto_stats.get('match_rate', 0):.1%}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
        
        # Manual categorization stats
        apply_stats = stats.get('apply_stats', {})
        if apply_stats:
            row += 1
            ttk.Label(stats_frame, text="Manual Categorization:", font=("TkDefaultFont", 10, "bold")).grid(
                row=row, column=0, sticky=tk.W, pady=(10, 5))
            row += 1
            
            ttk.Label(stats_frame, text=f"Rows updated: {apply_stats.get('rows_updated', 0)}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
            
            ttk.Label(stats_frame, text=f"Final coverage: {apply_stats.get('coverage_rate', 0):.1%}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
        
        # Dictionary update stats
        update_result = stats.get('update_result', {})
        if update_result:
            row += 1
            ttk.Label(stats_frame, text="Dictionary Updates:", font=("TkDefaultFont", 10, "bold")).grid(
                row=row, column=0, sticky=tk.W, pady=(10, 5))
            row += 1
            
            ttk.Label(stats_frame, text=f"New mappings added: {update_result.get('total_added', 0)}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
            
            ttk.Label(stats_frame, text=f"Conflicts found: {update_result.get('total_conflicts', 0)}").grid(
                row=row, column=0, sticky=tk.W)
            row += 1
        
        # Configure grid weights
        stats_frame.grid_columnconfigure(0, weight=1)
    
    def _show_review_step(self):
        """Show final review step"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Create review content
        review_frame = ttk.LabelFrame(self.content_frame, text="Final Review", padding="10")
        review_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=10, pady=10)
        
        # Instructions
        instructions = """
        Review the categorization results before finalizing.
        
        You can:
        - Review the statistics above
        - Export the categorized data
        - Modify categories if needed
        - Proceed to finalize the categorization
        """
        
        instruction_label = ttk.Label(review_frame, text=instructions, wraplength=500, justify=tk.LEFT)
        instruction_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(review_frame)
        button_frame.grid(row=1, column=0, pady=(10, 0))
        
        export_button = ttk.Button(button_frame, text="Export Categorized Data", 
                                  command=self._export_categorized_data)
        export_button.pack(side=tk.LEFT, padx=(0, 5))
        
        finalize_button = ttk.Button(button_frame, text="Finalize Categorization", 
                                    command=self._finalize_categorization)
        finalize_button.pack(side=tk.LEFT)
        
        # Configure grid weights
        review_frame.grid_columnconfigure(0, weight=1)
    
    def _open_excel_file(self):
        """Open the manual categorization Excel file"""
        if self.categorization_result and 'summary' in self.categorization_result:
            excel_path = self.categorization_result['summary'].get('manual_excel_generated', '')
            print(f"[DEBUG] UI trying to open: {excel_path}, exists: {os.path.exists(excel_path)}")
            if excel_path and os.path.exists(excel_path):
                try:
                    os.startfile(excel_path)  # Windows
                except AttributeError:
                    try:
                        import subprocess
                        subprocess.run(['open', excel_path])  # macOS
                    except FileNotFoundError:
                        try:
                            subprocess.run(['xdg-open', excel_path])  # Linux
                        except FileNotFoundError:
                            messagebox.showwarning("Warning", 
                                                 f"Could not open file automatically.\nPlease open: {excel_path}")
            else:
                messagebox.showerror("Error", "Excel file not found")
    
    def _upload_completed_file(self):
        """Upload the completed manual categorization file"""
        file_path = filedialog.askopenfilename(
            title="Select Completed Manual Categorization File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                from core.manual_categorizer import process_manual_categorizations, apply_manual_categories
                
                # Process the uploaded file
                manual_cats = process_manual_categorizations(Path(file_path))
                
                if manual_cats:
                    # Apply the manual categorizations
                    mapped_df = self._get_dataframe_from_mapping()
                    if mapped_df is not None:
                        apply_result = apply_manual_categories(mapped_df, manual_cats)
                        self.final_dataframe = apply_result['updated_dataframe']
                        
                        # Update the categorization result
                        if self.categorization_result:
                            self.categorization_result['all_stats']['apply_stats'] = apply_result['statistics']
                        
                        messagebox.showinfo("Success", 
                                          f"Successfully applied {len(manual_cats)} manual categorizations")
                        
                        # Move to next step
                        self._show_statistics_step()
                        self.next_button.config(state=tk.NORMAL)
                    else:
                        messagebox.showerror("Error", "No data available for categorization")
                else:
                    messagebox.showwarning("Warning", "No manual categorizations found in the file")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file: {str(e)}")
    
    def _export_categorized_data(self):
        """Export the categorized data"""
        if self.final_dataframe is None:
            messagebox.showwarning("Warning", "No categorized data available")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Categorized Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.final_dataframe.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    self.final_dataframe.to_excel(file_path, index=False)
                else:
                    self.final_dataframe.to_csv(file_path, index=False)
                
                messagebox.showinfo("Success", f"Data exported to: {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def _finalize_categorization(self):
        """Finalize the categorization process"""
        if self.final_dataframe is not None:
            # Call the completion callback
            if self.on_complete:
                self.on_complete(self.final_dataframe, self.categorization_result)
            
            messagebox.showinfo("Success", "Categorization completed successfully!")
            self.dialog.destroy()
        else:
            messagebox.showwarning("Warning", "No categorized data available")
    
    def _show_error(self, message):
        """Show error message"""
        messagebox.showerror("Categorization Error", message)
        self.dialog.destroy()
    
    def _show_success(self):
        """Show success and move to next step"""
        self._show_manual_categorization_step()
        self.next_button.config(state=tk.NORMAL)
    
    def _on_next(self):
        """Handle next button click"""
        # This would implement step navigation
        pass
    
    def _on_back(self):
        """Handle back button click"""
        # This would implement step navigation
        pass
    
    def _on_cancel(self):
        """Handle cancel button click"""
        if messagebox.askyesno("Cancel", "Are you sure you want to cancel the categorization?"):
            self.dialog.destroy()
    
    def _on_close(self):
        """Handle window close"""
        self._on_cancel()


def show_categorization_dialog(parent, controller, file_mapping, on_complete=None):
    """
    Show the categorization dialog
    
    Args:
        parent: Parent window
        controller: Main application controller
        file_mapping: File mapping object
        on_complete: Callback function when categorization is complete
    
    Returns:
        CategorizationDialog instance
    """
    dialog = CategorizationDialog(parent, controller, file_mapping, on_complete)
    return dialog 