"""
Progress Dialog for BOQ Tools
Shows real-time processing status and progress
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional, Callable
import platform
import time
import threading
from datetime import datetime, timedelta

# Color coding for status
def status_color(status: str) -> str:
    colors = {
        'running': '#2196F3',      # Blue
        'success': '#4CAF50',      # Green
        'error': '#F44336',        # Red
        'warning': '#FF9800',      # Orange
        'cancelled': '#757575'     # Gray
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


class ProgressDialog:
    def __init__(self, parent, title: str = "Processing", on_cancel: Optional[Callable] = None):
        """
        Initialize the progress dialog
        
        Args:
            parent: Parent window
            title: Dialog title
            on_cancel: Callback function when user cancels
        """
        self.parent = parent
        self.title = title
        self.on_cancel = on_cancel
        self.cancelled = False
        self.start_time = None
        self.current_step = 0
        self.total_steps = 0
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x400")
        self.dialog.minsize(400, 300)
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
        
        # Start time tracking
        self.start_time = time.time()
        
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
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Title
        title_label = ttk.Label(main_frame, text=self.title, font=("TkDefaultFont", 12, "bold"))
        title_label.pack(pady=(0, 15))
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Current Status", padding=10)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="Initializing...")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                     font=("TkDefaultFont", 10))
        self.status_label.pack(anchor=tk.W)
        
        # Step label
        self.step_var = tk.StringVar(value="")
        self.step_label = ttk.Label(status_frame, textvariable=self.step_var, 
                                   font=("TkDefaultFont", 9))
        self.step_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding=10)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=400, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        # Progress percentage
        self.progress_text_var = tk.StringVar(value="0%")
        progress_text = ttk.Label(progress_frame, textvariable=self.progress_text_var)
        progress_text.pack()
        
        # Time frame
        time_frame = ttk.LabelFrame(main_frame, text="Time Information", padding=10)
        time_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Time labels
        self.elapsed_var = tk.StringVar(value="Elapsed: 00:00")
        self.estimated_var = tk.StringVar(value="Estimated: --:--")
        self.remaining_var = tk.StringVar(value="Remaining: --:--")
        
        ttk.Label(time_frame, textvariable=self.elapsed_var).pack(anchor=tk.W)
        ttk.Label(time_frame, textvariable=self.estimated_var).pack(anchor=tk.W)
        ttk.Label(time_frame, textvariable=self.remaining_var).pack(anchor=tk.W)
        
        # Details frame
        details_frame = ttk.LabelFrame(main_frame, text="Processing Details", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Details text
        self.details_text = tk.Text(details_frame, height=8, wrap=tk.WORD, 
                                   font=("TkDefaultFont", 9))
        details_scroll = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, 
                                      command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=details_scroll.set)
        
        self.details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        details_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Error frame (initially hidden)
        self.error_frame = ttk.LabelFrame(main_frame, text="Errors", padding=10)
        self.error_text = tk.Text(self.error_frame, height=4, wrap=tk.WORD, 
                                 font=("TkDefaultFont", 9), background="#ffebee")
        error_scroll = ttk.Scrollbar(self.error_frame, orient=tk.VERTICAL, 
                                    command=self.error_text.yview)
        self.error_text.configure(yscrollcommand=error_scroll.set)
        
        self.error_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        error_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Buttons
        self.cancel_button = ttk.Button(button_frame, text="Cancel", command=self._on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=5)
        
        self.retry_button = ttk.Button(button_frame, text="Retry", command=self._on_retry, state=tk.DISABLED)
        self.retry_button.pack(side=tk.RIGHT, padx=5)
        
        self.details_button = ttk.Button(button_frame, text="Show Details", command=self._toggle_details)
        self.details_button.pack(side=tk.LEFT, padx=5)
        
        # Initially hide details
        self.details_visible = False
        details_frame.pack_forget()

    def _bind_shortcuts(self):
        """Bind keyboard shortcuts"""
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
        self.dialog.bind('<Control-c>', lambda e: self._on_cancel())
        self.dialog.bind('<F5>', lambda e: self._refresh_display())

    def set_total_steps(self, total: int):
        """Set the total number of steps"""
        self.total_steps = total
        self._update_step_display()

    def set_current_step(self, step: int, step_name: str = ""):
        """Set the current step"""
        self.current_step = step
        self.step_name = step_name
        self._update_step_display()
        self._update_progress()

    def set_status(self, status: str, color: str = None):
        """Set the status message"""
        self.status_var.set(status)
        if color:
            self.status_label.configure(foreground=color)
        self._add_detail(f"Status: {status}")

    def set_progress(self, percentage: float):
        """Set the progress percentage"""
        self.progress_var.set(percentage)
        self.progress_text_var.set(f"{percentage:.1f}%")
        self._update_time_estimates()

    def add_detail(self, message: str):
        """Add a detail message"""
        self._add_detail(message)

    def add_error(self, error: str):
        """Add an error message"""
        self._add_error(error)
        self._show_error_frame()

    def add_warning(self, warning: str):
        """Add a warning message"""
        self._add_detail(f"Warning: {warning}")

    def _update_step_display(self):
        """Update the step display"""
        if self.total_steps > 0:
            step_text = f"Step {self.current_step} of {self.total_steps}"
            if self.step_name:
                step_text += f": {self.step_name}"
            self.step_var.set(step_text)
        else:
            self.step_var.set(self.step_name if self.step_name else "")

    def _update_progress(self):
        """Update progress based on current step"""
        if self.total_steps > 0:
            percentage = (self.current_step / self.total_steps) * 100
            self.set_progress(percentage)

    def _update_time_estimates(self):
        """Update time estimates"""
        if not self.start_time:
            return
        
        elapsed = time.time() - self.start_time
        elapsed_str = str(timedelta(seconds=int(elapsed)))
        self.elapsed_var.set(f"Elapsed: {elapsed_str}")
        
        # Estimate remaining time
        if self.progress_var.get() > 0:
            total_estimated = elapsed / (self.progress_var.get() / 100)
            remaining = total_estimated - elapsed
            if remaining > 0:
                remaining_str = str(timedelta(seconds=int(remaining)))
                self.remaining_var.set(f"Remaining: {remaining_str}")
                
                total_str = str(timedelta(seconds=int(total_estimated)))
                self.estimated_var.set(f"Estimated Total: {total_str}")

    def _add_detail(self, message: str):
        """Add a detail message to the details text"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.details_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.details_text.see(tk.END)
        self.dialog.update_idletasks()

    def _add_error(self, error: str):
        """Add an error message to the error text"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.error_text.insert(tk.END, f"[{timestamp}] ERROR: {error}\n")
        self.error_text.see(tk.END)
        self.dialog.update_idletasks()

    def _show_error_frame(self):
        """Show the error frame"""
        if not self.error_frame.winfo_ismapped():
            self.error_frame.pack(fill=tk.X, pady=(0, 10), before=self.dialog.winfo_children()[-1])

    def _hide_error_frame(self):
        """Hide the error frame"""
        if self.error_frame.winfo_ismapped():
            self.error_frame.pack_forget()

    def _toggle_details(self):
        """Toggle the details frame visibility"""
        if self.details_visible:
            # Hide details
            for child in self.dialog.winfo_children():
                if isinstance(child, ttk.LabelFrame) and "Processing Details" in str(child):
                    child.pack_forget()
            self.details_button.configure(text="Show Details")
            self.details_visible = False
        else:
            # Show details
            for child in self.dialog.winfo_children():
                if isinstance(child, ttk.LabelFrame) and "Processing Details" in str(child):
                    child.pack(fill=tk.BOTH, expand=True, pady=(0, 10), before=self.dialog.winfo_children()[-1])
            self.details_button.configure(text="Hide Details")
            self.details_visible = True

    def _refresh_display(self):
        """Refresh the display"""
        self.dialog.update_idletasks()

    def _on_cancel(self):
        """Handle cancel button click"""
        if messagebox.askyesno("Cancel Processing", "Are you sure you want to cancel the current operation?"):
            self.cancelled = True
            self.set_status("Cancelled by user", status_color('cancelled'))
            self.cancel_button.configure(state=tk.DISABLED)
            self.retry_button.configure(state=tk.NORMAL)
            
            if self.on_cancel:
                self.on_cancel()

    def _on_retry(self):
        """Handle retry button click"""
        self.cancelled = False
        self.set_status("Retrying...", status_color('running'))
        self.cancel_button.configure(state=tk.NORMAL)
        self.retry_button.configure(state=tk.DISABLED)
        self._hide_error_frame()
        
        # Clear error text
        self.error_text.delete(1.0, tk.END)

    def set_completed(self, success: bool = True, message: str = ""):
        """Mark the operation as completed"""
        if success:
            self.set_status("Completed successfully", status_color('success'))
            self.set_progress(100.0)
        else:
            self.set_status("Completed with errors", status_color('error'))
        
        if message:
            self._add_detail(message)
        
        self.cancel_button.configure(state=tk.DISABLED)
        self.retry_button.configure(state=tk.NORMAL)

    def is_cancelled(self) -> bool:
        """Check if the operation was cancelled"""
        return self.cancelled

    def close(self):
        """Close the dialog"""
        self.dialog.destroy()

    def show(self):
        """Show the dialog"""
        self.dialog.wait_window()
        return not self.cancelled


# Convenience function for simple progress display
def show_progress_dialog(parent, title: str = "Processing", on_cancel: Optional[Callable] = None) -> ProgressDialog:
    """
    Show a progress dialog
    
    Args:
        parent: Parent window
        title: Dialog title
        on_cancel: Optional callback function
        
    Returns:
        ProgressDialog instance
    """
    return ProgressDialog(parent, title, on_cancel)


# Example usage with threading
class ProcessingThread(threading.Thread):
    def __init__(self, progress_dialog: ProgressDialog, processing_func: Callable):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.processing_func = processing_func
        self.daemon = True

    def run(self):
        try:
            self.processing_func(self.progress_dialog)
        except Exception as e:
            self.progress_dialog.add_error(str(e))
            self.progress_dialog.set_completed(False, f"Processing failed: {e}")


def run_with_progress(parent, title: str, processing_func: Callable, on_complete: Optional[Callable] = None):
    """
    Run a processing function with progress dialog
    
    Args:
        parent: Parent window
        title: Dialog title
        processing_func: Function to run (should accept ProgressDialog as parameter)
        on_complete: Optional callback when processing completes
    """
    progress_dialog = ProgressDialog(parent, title)
    
    def on_cancel():
        # Handle cancellation
        pass
    
    progress_dialog.on_cancel = on_cancel
    
    # Start processing in background thread
    thread = ProcessingThread(progress_dialog, processing_func)
    thread.start()
    
    # Show dialog
    success = progress_dialog.show()
    
    if on_complete:
        on_complete(success) 