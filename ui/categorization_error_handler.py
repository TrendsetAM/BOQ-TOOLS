"""
Categorization Error Handler for BOQ Tools
Provides comprehensive error handling for categorization UI components
"""

import tkinter as tk
from tkinter import messagebox
import logging
import traceback
from typing import Dict, Any, Optional, Callable
from pathlib import Path
import sys

logger = logging.getLogger(__name__)


class CategorizationErrorHandler:
    """Handles errors in categorization UI components"""
    
    def __init__(self, parent_window=None):
        """
        Initialize the error handler
        
        Args:
            parent_window: Parent window for error dialogs
        """
        self.parent_window = parent_window
        self.error_count = 0
        self.max_errors = 5  # Maximum errors before stopping
        self.error_log = []
    
    def handle_error(self, error: Exception, context: str = "", 
                    show_dialog: bool = True, log_error: bool = True) -> bool:
        """
        Handle an error in the categorization process
        
        Args:
            error: The exception that occurred
            context: Context information about where the error occurred
            show_dialog: Whether to show an error dialog
            log_error: Whether to log the error
            
        Returns:
            True if the error was handled successfully, False otherwise
        """
        self.error_count += 1
        
        # Log the error
        if log_error:
            self._log_error(error, context)
        
        # Add to error log
        self.error_log.append({
            'error': str(error),
            'context': context,
            'traceback': traceback.format_exc(),
            'count': self.error_count
        })
        
        # Check if we've exceeded max errors
        if self.error_count >= self.max_errors:
            self._handle_max_errors_reached()
            return False
        
        # Show error dialog if requested
        if show_dialog:
            self._show_error_dialog(error, context)
        
        return True
    
    def _log_error(self, error: Exception, context: str):
        """Log the error with context"""
        error_msg = f"Categorization Error in {context}: {str(error)}"
        logger.error(error_msg, exc_info=True)
    
    def _show_error_dialog(self, error: Exception, context: str):
        """Show an error dialog to the user"""
        if not self.parent_window:
            return
        
        # Create error dialog
        dialog = tk.Toplevel(self.parent_window)
        dialog.title("Categorization Error")
        dialog.geometry("600x400")
        dialog.transient(self.parent_window)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"600x400+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Error icon and title
        title_frame = tk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        
        error_label = tk.Label(title_frame, text="⚠️", font=("TkDefaultFont", 24))
        error_label.pack(side=tk.LEFT, padx=(0, 10))
        
        title_label = tk.Label(title_frame, text="Categorization Error", 
                              font=("TkDefaultFont", 14, "bold"))
        title_label.pack(side=tk.LEFT)
        
        # Context information
        if context:
            context_label = tk.Label(main_frame, text=f"Context: {context}", 
                                   font=("TkDefaultFont", 10, "bold"))
            context_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Error message
        error_frame = tk.LabelFrame(main_frame, text="Error Details", padx=10, pady=10)
        error_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Error text with scrollbar
        text_frame = tk.Frame(error_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        error_text = tk.Text(text_frame, wrap=tk.WORD, height=10)
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=error_text.yview)
        error_text.configure(yscrollcommand=scrollbar.set)
        
        error_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Insert error information
        error_text.insert(tk.END, f"Error: {str(error)}\n\n")
        error_text.insert(tk.END, f"Error Count: {self.error_count}/{self.max_errors}\n\n")
        
        # Add traceback if available
        try:
            tb = traceback.format_exc()
            if tb and tb != "NoneType: None\n":
                error_text.insert(tk.END, "Traceback:\n")
                error_text.insert(tk.END, tb)
        except:
            pass
        
        error_text.config(state=tk.DISABLED)
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Copy error button
        copy_btn = tk.Button(button_frame, text="Copy Error Details", 
                           command=lambda: self._copy_error_to_clipboard(error, context))
        copy_btn.pack(side=tk.LEFT)
        
        # Continue button
        continue_btn = tk.Button(button_frame, text="Continue", 
                               command=dialog.destroy)
        continue_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Retry button (if applicable)
        retry_btn = tk.Button(button_frame, text="Retry", 
                            command=lambda: self._retry_operation(dialog))
        retry_btn.pack(side=tk.RIGHT)
        
        # Configure button frame
        button_frame.grid_columnconfigure(0, weight=1)
    
    def _handle_max_errors_reached(self):
        """Handle when maximum errors are reached"""
        if not self.parent_window:
            return
        
        messagebox.showerror(
            "Too Many Errors",
            f"Too many errors have occurred ({self.error_count}).\n"
            "The categorization process will be stopped.\n\n"
            "Please check the error log and try again."
        )
    
    def _copy_error_to_clipboard(self, error: Exception, context: str):
        """Copy error details to clipboard"""
        try:
            error_text = f"Context: {context}\n"
            error_text += f"Error: {str(error)}\n\n"
            error_text += f"Traceback:\n{traceback.format_exc()}"
            
            self.parent_window.clipboard_clear()
            self.parent_window.clipboard_append(error_text)
            
            messagebox.showinfo("Copied", "Error details copied to clipboard")
        except Exception as e:
            logger.error(f"Failed to copy error to clipboard: {e}")
    
    def _retry_operation(self, dialog):
        """Retry the failed operation"""
        dialog.destroy()
        # This would be implemented by the calling code
        # For now, just close the dialog
    
    def get_error_summary(self) -> Dict[str, Any]:
        """Get a summary of all errors"""
        return {
            'total_errors': self.error_count,
            'max_errors': self.max_errors,
            'errors': self.error_log
        }
    
    def reset_error_count(self):
        """Reset the error count"""
        self.error_count = 0
        self.error_log.clear()
    
    def export_error_log(self, file_path: Path):
        """Export error log to file"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("BOQ Tools - Categorization Error Log\n")
                f.write("=" * 50 + "\n\n")
                
                for error_info in self.error_log:
                    f.write(f"Error #{error_info['count']}\n")
                    f.write(f"Context: {error_info['context']}\n")
                    f.write(f"Error: {error_info['error']}\n")
                    f.write(f"Traceback:\n{error_info['traceback']}\n")
                    f.write("-" * 30 + "\n\n")
            
            return True
        except Exception as e:
            logger.error(f"Failed to export error log: {e}")
            return False


class CategorizationValidationError(Exception):
    """Custom exception for categorization validation errors"""
    
    def __init__(self, message: str, field: str = "", value: Any = None):
        self.message = message
        self.field = field
        self.value = value
        super().__init__(self.message)


class CategorizationTimeoutError(Exception):
    """Custom exception for categorization timeout errors"""
    
    def __init__(self, operation: str, timeout_seconds: int):
        self.operation = operation
        self.timeout_seconds = timeout_seconds
        self.message = f"Operation '{operation}' timed out after {timeout_seconds} seconds"
        super().__init__(self.message)


def validate_categorization_input(dataframe, required_columns=None):
    """
    Validate input for categorization
    
    Args:
        dataframe: DataFrame to validate
        required_columns: List of required columns
    
    Raises:
        CategorizationValidationError: If validation fails
    """
    if dataframe is None:
        raise CategorizationValidationError("DataFrame is None")
    
    if len(dataframe) == 0:
        raise CategorizationValidationError("DataFrame is empty")
    
    if required_columns:
        missing_columns = [col for col in required_columns if col not in dataframe.columns]
        if missing_columns:
            raise CategorizationValidationError(
                f"Missing required columns: {missing_columns}",
                field="columns",
                value=missing_columns
            )
    
    # Check for required 'Description' column for categorization
    if 'Description' not in dataframe.columns:
        raise CategorizationValidationError(
            "Missing 'Description' column required for categorization",
            field="Description",
            value=None
        )


def handle_categorization_timeout(operation: str, timeout_seconds: int = 30):
    """
    Decorator to handle timeouts in categorization operations
    
    Args:
        operation: Name of the operation
        timeout_seconds: Timeout in seconds
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            import signal
            
            def timeout_handler(signum, frame):
                raise CategorizationTimeoutError(operation, timeout_seconds)
            
            # Set up timeout handler
            old_handler = signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout_seconds)
            
            try:
                result = func(*args, **kwargs)
                signal.alarm(0)  # Cancel alarm
                return result
            except CategorizationTimeoutError:
                raise
            except Exception as e:
                signal.alarm(0)  # Cancel alarm
                raise e
            finally:
                signal.signal(signal.SIGALRM, old_handler)
        
        return wrapper
    return decorator


def safe_categorization_operation(operation_func: Callable, error_handler: CategorizationErrorHandler,
                                context: str = "", *args, **kwargs):
    """
    Safely execute a categorization operation with error handling
    
    Args:
        operation_func: Function to execute
        error_handler: Error handler instance
        context: Context for error reporting
        *args, **kwargs: Arguments for the operation function
    
    Returns:
        Result of the operation or None if failed
    """
    try:
        return operation_func(*args, **kwargs)
    except CategorizationValidationError as e:
        error_handler.handle_error(e, f"{context} - Validation Error")
        return None
    except CategorizationTimeoutError as e:
        error_handler.handle_error(e, f"{context} - Timeout Error")
        return None
    except Exception as e:
        error_handler.handle_error(e, f"{context} - Unexpected Error")
        return None


# Global error handler instance
_global_error_handler = None


def get_global_error_handler(parent_window=None) -> CategorizationErrorHandler:
    """Get or create the global error handler"""
    global _global_error_handler
    if _global_error_handler is None:
        _global_error_handler = CategorizationErrorHandler(parent_window)
    return _global_error_handler


def set_global_error_handler(error_handler: CategorizationErrorHandler):
    """Set the global error handler"""
    global _global_error_handler
    _global_error_handler = error_handler 