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
import pandas as pd

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
        
        # Make dialog modal without enforcing a global grab
        self.dialog.transient(self.parent)
        self.dialog.focus_set()
        self.dialog.lift()
        
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
            self.current_step.set("Step 1/4: Loading category dictionary")
        elif percent < 40:
            self.current_step.set("Step 2/4: Auto-categorizing rows")
        elif percent < 60:
            self.current_step.set("Step 3/4: Collecting unmatched descriptions")
        elif percent < 100:
            self.current_step.set("Step 4/4: Generating manual categorization file")
        else:
            self.current_step.set("Ready for manual categorization")
    
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
        
        # Add buttons for navigation
        button_frame = ttk.Frame(stats_frame)
        button_frame.grid(row=row+1, column=0, pady=(20, 0))
        
        stats_button = ttk.Button(button_frame, text="View Statistics", 
                                 command=self._show_statistics_step)
        stats_button.pack(side=tk.LEFT, padx=(0, 5))
        
        export_button = ttk.Button(button_frame, text="Export Data", 
                                  command=self._export_categorized_data)
        export_button.pack(side=tk.LEFT, padx=(0, 5))
        
        finalize_button = ttk.Button(button_frame, text="Finalize", 
                                    command=self._finalize_categorization)
        finalize_button.pack(side=tk.LEFT)
    
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
            # print(f"[DEBUG] UI trying to open: {excel_path}, exists: {os.path.exists(excel_path)}")
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
                from core.manual_categorizer import process_manual_categorizations, apply_manual_categories, update_master_dictionary
                from core.category_dictionary import CategoryDictionary
                
                # Process the uploaded file
                manual_cats = process_manual_categorizations(Path(file_path))
                
                if manual_cats:
                    # Get the auto-categorized DataFrame from the categorization result
                    if self.categorization_result and 'final_dataframe' in self.categorization_result:
                        auto_df = self.categorization_result['final_dataframe']
                        
                        # Apply the manual categorizations to the auto-categorized DataFrame
                        apply_result = apply_manual_categories(auto_df, manual_cats)
                        self.final_dataframe = apply_result['updated_dataframe']
                        
                        # Update the categorization result
                        if self.categorization_result:
                            self.categorization_result['all_stats']['apply_stats'] = apply_result['statistics']
                            self.categorization_result['all_stats']['manual_categorization_count'] = len(manual_cats)
                        
                        messagebox.showinfo("Success", 
                                          f"Successfully applied {len(manual_cats)} manual categorizations")
                        
                        # Prompt user to update dictionary
                        self._prompt_dictionary_update(manual_cats)
                        
                        # Delete the temporary Excel file
                        self._cleanup_temp_files()
                        
                        # Move to statistics step
                        self._show_statistics_step()
                        self.next_button.config(state=tk.NORMAL)
                    else:
                        messagebox.showerror("Error", "No auto-categorized data available")
                else:
                    messagebox.showwarning("Warning", "No manual categorizations found in the file")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file: {str(e)}")
    
    def _prompt_dictionary_update(self, manual_cats):
        """Prompt user to update the dictionary with new manual categorizations"""
        response = messagebox.askyesno(
            "Update Dictionary", 
            f"Would you like to update the category dictionary with {len(manual_cats)} new manual categorizations?\n\n"
            "This will add the new description-category mappings to the master dictionary for future auto-categorization."
        )
        
        if response:
            try:
                from core.manual_categorizer import update_master_dictionary
                from core.category_dictionary import CategoryDictionary
                from pathlib import Path
                
                # Load the category dictionary (uses user config path automatically)
                category_dict = CategoryDictionary()
                
                # Update the dictionary
                update_result = update_master_dictionary(category_dict, manual_cats)
                
                # Update the categorization result with dictionary update info
                if self.categorization_result:
                    self.categorization_result['all_stats']['update_result'] = update_result
                
                # Show update results
                message = f"Dictionary updated successfully!\n\n"
                message += f"New mappings added: {update_result.get('total_added', 0)}\n"
                message += f"Conflicts found: {update_result.get('total_conflicts', 0)}\n"
                message += f"Already existing: {update_result.get('total_skipped', 0)}\n"
                
                if update_result.get('backup_path'):
                    message += f"\nBackup created at: {update_result['backup_path']}"
                
                messagebox.showinfo("Dictionary Updated", message)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update dictionary: {str(e)}")
    
    def _cleanup_temp_files(self):
        """Clean up temporary files after categorization is complete"""
        try:
            if self.categorization_result and 'summary' in self.categorization_result:
                excel_path = self.categorization_result['summary'].get('manual_excel_generated', '')
                if excel_path and os.path.exists(excel_path):
                    os.remove(excel_path)
                    # print(f"[DEBUG] Deleted temporary Excel file: {excel_path}")
        except Exception as e:
            print(f"[WARNING] Could not delete temporary file: {e}")
    
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
        # print("[DEBUG] Dialog finalize called. on_complete:", self.on_complete)
        if self.final_dataframe is not None:
            # Call the completion callback to update the main window
            if self.on_complete:
                # print("[DEBUG] Calling on_complete from dialog with final_dataframe:", type(self.final_dataframe), "categorization_result:", type(self.categorization_result))
                self.on_complete(self.final_dataframe, self.categorization_result)
            self.dialog.destroy()
        else:
            messagebox.showwarning("Warning", "No categorized data available")
    
    def _show_error(self, message):
        """Show error message"""
        messagebox.showerror("Categorization Error", message)
        self.dialog.destroy()
    
    def _show_success(self):
        """Show success and move to next step"""
        # Check if manual categorization is needed
        if self._needs_manual_categorization():
            self._show_manual_categorization_step()
        else:
            # All rows were auto-categorized successfully
            self._show_completion_step()
        self.next_button.config(state=tk.NORMAL)
    
    def _needs_manual_categorization(self):
        """Check if manual categorization is needed based on unmatched rows"""
        if not self.categorization_result:
            return False
        
        # Check if there are unmatched rows
        stats = self.categorization_result.get('all_stats', {})
        auto_stats = stats.get('auto_stats', {})
        
        total_rows = auto_stats.get('total_rows', 0)
        matched_rows = auto_stats.get('matched_rows', 0)
        unmatched_rows = total_rows - matched_rows
        
        # Also check if a manual Excel file was generated
        has_manual_excel = self.categorization_result.get('summary', {}).get('manual_excel_generated', '')
        
        return unmatched_rows > 0 and has_manual_excel
    
    def _show_completion_step(self):
        """Show completion step when all rows were auto-categorized"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Hide navigation buttons in the completion step
        if hasattr(self, 'back_button') and self.back_button:
            self.back_button.pack_forget()
        if hasattr(self, 'next_button') and self.next_button:
            self.next_button.pack_forget()
        if hasattr(self, 'cancel_button') and self.cancel_button:
            self.cancel_button.pack_forget()
        
        # Create completion content
        completion_frame = ttk.LabelFrame(self.content_frame, text="Categorization Complete", padding="10")
        completion_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=10, pady=10)
        
        # Success message
        message = """
        🎉 Excellent! All rows were automatically categorized successfully.
        
        No manual categorization was needed - the system was able to match all descriptions to categories using the existing dictionary.
        
        You can now proceed to review and export the final categorized data in the main window.
        """
        
        message_label = ttk.Label(completion_frame, text=message, wraplength=500, justify=tk.LEFT)
        message_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(completion_frame)
        button_frame.grid(row=1, column=0, pady=(10, 0))
        
        next_button = ttk.Button(button_frame, text="Next", command=self._finalize_categorization)
        next_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # Configure grid weights
        completion_frame.grid_columnconfigure(0, weight=1)
    
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