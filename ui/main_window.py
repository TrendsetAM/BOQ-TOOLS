"""
BOQ Tools Main Window
Comprehensive GUI for Excel file processing and BOQ analysis
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import platform
from pathlib import Path
from typing import List, Dict, Any
import logging
import threading

# Get a logger for this module
logger = logging.getLogger(__name__)

# Optional imports
try:
    from ttkthemes import ThemedTk
    THEME_AVAILABLE = True
except ImportError:
    THEME_AVAILABLE = False

try:
    import tkinterdnd2 as TkinterDnD
    DND_AVAILABLE = True
    DND_FILES = TkinterDnD.DND_FILES
except ImportError:
    DND_AVAILABLE = False

# Import settings dialog
try:
    from ui.settings_dialog import show_settings_dialog
    SETTINGS_AVAILABLE = True
except ImportError:
    SETTINGS_AVAILABLE = False

# Color coding for confidence
def confidence_color(score):
    if score >= 0.8:
        return '#4CAF50'  # Green
    elif score >= 0.6:
        return '#FFC107'  # Amber
    else:
        return '#F44336'  # Red


def tooltip(widget, text):
    """Simple tooltip for a widget"""
    def on_enter(event):
        widget._tip = tk.Toplevel()
        widget._tip.wm_overrideredirect(True)
        x = event.x_root + 10
        y = event.y_root + 10
        widget._tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(widget._tip, text=text, background="#ffffe0", relief='solid', borderwidth=1, font=("TkDefaultFont", 9))
        label.pack()
    def on_leave(event):
        if hasattr(widget, '_tip') and widget._tip is not None:
            widget._tip.destroy()
            widget._tip = None
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)


class MainWindow:
    def __init__(self, controller, root=None):
        self.controller = controller
        self.root = root or self._create_root()
        self.root.title("BOQ Tools - Excel Processor")
        self.root.geometry("1200x700")
        self.root.minsize(900, 600)
        self._setup_style()
        self._create_menu()
        # Initialize variables before creating widgets
        self.open_files = {}
        self.status_var = tk.StringVar(value="Ready.")
        self.progress_var = tk.DoubleVar(value=0)
        self.status_bar_visible = True
        self.file_mapping = None  # To store the full file mapping object
        self.sheet_treeviews = {} # To store treeview widgets for each sheet
        # Create widgets
        self._create_main_widgets()
        self._setup_drag_and_drop()
        self._bind_shortcuts()
        self._update_status("Welcome to BOQ Tools!")

    def _create_root(self):
        if THEME_AVAILABLE:
            return ThemedTk(theme="arc")
        elif DND_AVAILABLE:
            return TkinterDnD.Tk()
        else:
            return tk.Tk()

    def _setup_style(self):
        style = ttk.Style(self.root)
        if platform.system() == 'Windows':
            style.theme_use('vista')
        elif platform.system() == 'Darwin':
            style.theme_use('aqua')
        else:
            style.theme_use('clam')
        style.configure('TNotebook.Tab', padding=[10, 5])
        style.configure('Treeview', rowheight=28)

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        # File
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_command(label="Export", accelerator="Ctrl+E", command=self.export_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", accelerator="Ctrl+Q", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        # Edit
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self._not_implemented)
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self._not_implemented)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        # View
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Status Bar", command=self._toggle_status_bar)
        menubar.add_cascade(label="View", menu=view_menu)
        # Tools
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Settings", command=self.open_settings)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        self.root.config(menu=menubar)

    def _create_main_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        # Top: Drag-and-drop zone
        self.drop_zone = ttk.Label(main_frame, text="Drop Excel files here or use File > Open", anchor="center", relief="ridge", padding=20)
        self.drop_zone.pack(fill=tk.X, padx=10, pady=8)
        tooltip(self.drop_zone, "Drag and drop .xlsx or .xls files here to open.")
        # Center: Tabbed interface for files
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        # Bottom: Status bar and progress
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w")
        self.status_label.pack(side=tk.LEFT, padx=8)
        self.progress = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100, length=180)
        self.progress.pack(side=tk.RIGHT, padx=8)

    def _setup_drag_and_drop(self):
        if DND_AVAILABLE:
            try:
                if hasattr(self.root, 'drop_target_register'):
                    self.root.drop_target_register(DND_FILES)
                    self.root.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass
        else:
            # Fallback: Make drop zone clickable
            self.drop_zone.bind('<Button-1>', lambda e: self.open_file())
            self.drop_zone.config(text="Click here to open Excel files or use File > Open")

    def _bind_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-e>', lambda e: self.export_file())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<Control-z>', lambda e: self._not_implemented())
        self.root.bind('<Control-y>', lambda e: self._not_implemented())

    def _update_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _toggle_status_bar(self):
        if self.status_bar_visible:
            self.status_label.pack_forget()
            self.progress.pack_forget()
        else:
            self.status_label.pack(side=tk.LEFT, padx=8)
            self.progress.pack(side=tk.RIGHT, padx=8)
        self.status_bar_visible = not self.status_bar_visible

    def _not_implemented(self):
        messagebox.showinfo("Not Implemented", "This feature is not implemented yet.")

    def open_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filenames = filedialog.askopenfilenames(title="Open Excel File", filetypes=filetypes)
        for file in filenames:
            self._open_excel_file(file)

    def _on_drop(self, event):
        if DND_AVAILABLE:
            files = self.root.tk.splitlist(event.data)
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    self._open_excel_file(file)
                else:
                    self._update_status(f"Unsupported file: {file}")

    def _open_excel_file(self, filepath):
        """Handle the file processing workflow"""
        if not filepath:
            return

        # Clear previous results, including treeview references
        self.sheet_treeviews.clear()
        if self.notebook:
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)

        self._update_status(f"Processing {os.path.basename(filepath)}, please wait...")
        self.progress_var.set(0)

        # Create a new tab for the file
        filename = os.path.basename(filepath)
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=filename)
        self.notebook.select(tab)

        loading_label = ttk.Label(tab, text="Analyzing file...")
        loading_label.pack(pady=40, padx=100)
        self.root.update_idletasks()

        def process_in_thread():
            """Runs the file processing in a background thread."""
            try:
                # Use the controller to process the file, sending progress updates to the UI thread
                self.file_mapping = self.controller.process_file(
                    Path(filepath), 
                    progress_callback=lambda p, m: self.root.after(0, self.update_progress, p, m)
                )
                # When done, update the UI from the main thread
                self.root.after(0, self._on_processing_complete, tab, filepath, self.file_mapping, loading_label)
            except Exception as e:
                logger.error(f"Failed to process file {filepath}: {e}", exc_info=True)
                self.root.after(0, self._on_processing_error, tab, filename, loading_label)

        threading.Thread(target=process_in_thread, daemon=True).start()

    def update_progress(self, percentage, message):
        """Thread-safe method to update the progress bar and status label."""
        self.progress_var.set(percentage)
        self._update_status(message)

    def _on_processing_complete(self, tab, filepath, file_mapping, loading_widget):
        """Callback for when file processing succeeds. Runs in the main UI thread."""
        loading_widget.destroy()
        # Store the real mapping data with the tab widget itself for later access (like exporting)
        setattr(tab, 'file_mapping', file_mapping)
        self._populate_file_tab(tab, file_mapping)
        self._update_status(f"Successfully processed: {os.path.basename(filepath)}")
        self.progress_var.set(100)
    
    def _on_processing_error(self, tab, filename, loading_widget):
        """Callback for when file processing fails. Runs in the main UI thread."""
        loading_widget.destroy()
        ttk.Label(tab, text=f"Failed to process {filename}.\nSee logs for details.", foreground="red").pack(pady=40)
        self._update_status(f"Error processing {filename}")
        self.progress_var.set(0)

    def _populate_file_tab(self, tab, file_mapping):
        """Populates a tab with the processed data from a file mapping."""
        # Clear any existing widgets (like loading/error labels)
        for widget in tab.winfo_children():
            widget.destroy()

        # Main frame for the tab content
        tab_frame = ttk.Frame(tab)
        tab_frame.pack(fill=tk.BOTH, expand=True)
        
        # Check if we have any sheets
        if not file_mapping.sheets:
            ttk.Label(tab_frame, text="No processable data found in this file.").pack(pady=40)
            return

        # Add export button at the top
        export_frame = ttk.Frame(tab_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        export_btn = ttk.Button(export_frame, text="Export Data", command=self.export_file)
        export_btn.pack(side=tk.RIGHT, padx=5)
        
        # Add global summary
        global_summary = ttk.LabelFrame(tab_frame, text="File Summary")
        global_summary.pack(fill=tk.X, padx=5, pady=5)
        
        global_text = f"""
Total Sheets: {len(file_mapping.sheets)}
Global Confidence: {file_mapping.global_confidence:.1%}
Export Ready: {'Yes' if file_mapping.export_ready else 'No'}
Processing Status: {file_mapping.processing_summary.successful_sheets} successful, {file_mapping.processing_summary.partial_sheets} partial
        """
        ttk.Label(global_summary, text=global_text, justify=tk.LEFT).pack(padx=10, pady=5)

        # Create a notebook for multiple sheets
        sheet_notebook = ttk.Notebook(tab_frame)
        sheet_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        for i, sheet_mapping in enumerate(file_mapping.sheets):
            sheet_frame = ttk.Frame(self.notebook)
            self.notebook.add(sheet_frame, text=sheet_mapping.sheet_name)

            # --- Configure grid layout for the sheet frame ---
            sheet_frame.grid_rowconfigure(0, weight=1)  # Make the canvas row expandable
            sheet_frame.grid_columnconfigure(0, weight=1)

            # --- Action Bar at the Bottom ---
            action_bar = ttk.Frame(sheet_frame, padding=5)
            action_bar.grid(row=1, column=0, sticky="ew") # Placed at the bottom

            propagate_btn = ttk.Button(
                action_bar,
                text="Apply These Mappings to All Other Sheets",
                command=lambda sm=sheet_mapping: self._propagate_mappings_for_sheet(sm)
            )
            propagate_btn.pack(side=tk.LEFT, padx=5)

            # --- Scrollable Main Area ---
            canvas = tk.Canvas(sheet_frame, highlightthickness=0)
            canvas.grid(row=0, column=0, sticky="nsew")

            scrollbar = ttk.Scrollbar(sheet_frame, orient="vertical", command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")
            
            canvas.configure(yscrollcommand=scrollbar.set)
            
            scrollable_frame = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            # --- Content inside the scrollable_frame ---
            # Sheet Summary Frame
            sheet_summary_frame = ttk.LabelFrame(scrollable_frame, text="Sheet Summary")
            sheet_summary_frame.pack(fill=tk.X, padx=10, pady=5, anchor="n")
            
            summary_text = (
                f"Processing Status: {sheet_mapping.processing_status.value}\n"
                f"Confidence: {sheet_mapping.overall_confidence:.1%}\n"
                f"Rows: {sheet_mapping.row_count}\n"
                f"Columns: {sheet_mapping.column_count}\n"
                f"Validation Score: {sheet_mapping.validation_summary.overall_score:.1%}"
            )
            ttk.Label(sheet_summary_frame, text=summary_text, justify=tk.LEFT).pack(padx=10, pady=5)
            
            # Column Mappings Frame
            col_frame = ttk.LabelFrame(scrollable_frame, text="Column Mappings (Double-click to edit) - Required columns are highlighted")
            col_frame.pack(fill=tk.X, padx=10, pady=5, anchor="n")
            
            col_tree = ttk.Treeview(col_frame, columns=('Column', 'Type', 'Confidence', 'Required', 'Actions'), show='headings', height=10)
            self.sheet_treeviews[sheet_mapping.sheet_name] = col_tree
            
            # Define column headings and styles
            col_tree.heading('Column', text='Column')
            col_tree.column('Column', width=200, anchor=tk.W)
            col_tree.heading('Type', text='Type')
            col_tree.heading('Confidence', text='Confidence')
            col_tree.heading('Required', text='Required')
            col_tree.heading('Actions', text='Actions')
            
            col_tree.column('Type', width=100, anchor=tk.W)
            col_tree.column('Confidence', width=100, anchor=tk.CENTER)
            col_tree.column('Required', width=80, anchor=tk.CENTER)
            col_tree.column('Actions', width=120, anchor=tk.CENTER)

            # Style tags
            col_tree.tag_configure('required', background='#E0E8F0', foreground='black')
            col_tree.tag_configure('edited', foreground='blue')

            # Populate column mappings
            if sheet_mapping and sheet_mapping.column_mappings:
                for col_mapping in sheet_mapping.column_mappings:
                    tags = []
                    if col_mapping.is_required:
                        tags.append('required')
                    if col_mapping.is_user_edited:
                        tags.append('edited')
                    
                    # Determine required text
                    if col_mapping.mapped_type == "ignore":
                        required_text = "Ignore"
                    elif col_mapping.is_required:
                        required_text = "Yes"
                    else:
                        required_text = "No"
                    
                    # Determine actions text
                    if col_mapping.is_user_edited:
                        actions_text = "Edited"
                    else:
                        actions_text = "Auto-detected"
                    
                    col_tree.insert('', 'end', values=(
                        col_mapping.original_header,
                        col_mapping.mapped_type,
                        f"{col_mapping.confidence:.1%}",
                        required_text,
                        actions_text
                    ), tags=tuple(tags))
            
            col_tree.bind('<Double-1>', lambda e, tree=col_tree, sheet=sheet_mapping: self._edit_column_mapping(tree, sheet))
            
            col_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Legend
            legend_text = "Legend: Required columns (Description, Quantity, Unit Price, Total Price, Unit, Code) are highlighted and essential for BOQ processing. Actions shows 'Auto-detected' for app decisions and 'Edited' for user changes."
            legend_label = ttk.Label(col_frame, text=legend_text, wraplength=800, justify=tk.LEFT, font=("Arial", 8))
            legend_label.pack(fill=tk.X, padx=5, pady=5)

            # Row Classifications Frame
            row_frame = ttk.LabelFrame(scrollable_frame, text="Row Classifications Summary")
            row_frame.pack(fill=tk.X, padx=10, pady=5, anchor="n")
            
            row_tree = ttk.Treeview(row_frame, columns=('Type', 'Count', 'Avg Confidence', 'Details'), show='headings', height=6)
            row_tree.heading('Type', text='Row Type')
            row_tree.column('Type', width=150, anchor=tk.W)
            row_tree.heading('Count', text='Count')
            row_tree.column('Count', width=80, anchor=tk.CENTER)
            row_tree.heading('Avg Confidence', text='Avg Confidence')
            row_tree.column('Avg Confidence', width=120, anchor=tk.CENTER)
            row_tree.heading('Details', text='Details')
            row_tree.column('Details', width=100, anchor=tk.CENTER)

            # Populate row classifications
            if sheet_mapping and sheet_mapping.row_classifications:
                summary = {}
                for row_class in sheet_mapping.row_classifications:
                    if row_class.row_type not in summary:
                        summary[row_class.row_type] = {'count': 0, 'total_confidence': 0}
                    summary[row_class.row_type]['count'] += 1
                    summary[row_class.row_type]['total_confidence'] += row_class.confidence

                sorted_summary = sorted(summary.items(), key=lambda item: item[1]['count'], reverse=True)
                
                for row_type, data in sorted_summary:
                    count = data['count']
                    avg_confidence = data['total_confidence'] / count if count > 0 else 0
                    
                    row_tree.insert('', tk.END, values=[row_type, count, f"{avg_confidence:.1%}", "View Details"])
                
                row_tree.bind('<Double-1>', lambda e, tree=row_tree, sheet=sheet_mapping: self._show_row_details(tree, sheet))
            
            row_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

            # Validation Issues Frame
            if sheet_mapping.validation_summary.error_count > 0 or sheet_mapping.validation_summary.warning_count > 0:
                val_frame = ttk.LabelFrame(scrollable_frame, text="Validation Issues")
                val_frame.pack(fill=tk.X, padx=10, pady=5, anchor="n")
                
                val_text = f"Errors: {sheet_mapping.validation_summary.error_count}\nWarnings: {sheet_mapping.validation_summary.warning_count}"
                ttk.Label(val_frame, text=val_text, foreground="red" if sheet_mapping.validation_summary.error_count > 0 else "orange").pack(padx=10, pady=5)

        self._update_status(f"Displaying data for {len(file_mapping.sheets)} sheets")

    def _propagate_mappings_for_sheet(self, source_sheet_mapping):
        """
        Finds all user-edited mappings on the source sheet and applies them
        to all other sheets in the workbook.
        """
        if not self.file_mapping:
            return

        edited_mappings = [
            cm for cm in source_sheet_mapping.column_mappings if cm.is_user_edited
        ]

        if not edited_mappings:
            messagebox.showinfo("No Changes", "No manually edited columns to propagate on this sheet.")
            return

        # Confirmation dialog
        msg = (
            f"This will apply {len(edited_mappings)} manually-edited column mappings from sheet "
            f"'{source_sheet_mapping.sheet_name}' to all other sheets.\n\n"
            "This may override existing mappings on other sheets. Are you sure you want to continue?"
        )
        if not messagebox.askyesno("Confirm Propagation", msg, icon='warning'):
            return

        # --- Propagation Logic ---
        propagated_count = 0
        changed_sheets = set()

        for target_sheet in self.file_mapping.sheets:
            if target_sheet.sheet_name == source_sheet_mapping.sheet_name:
                continue

            sheet_was_changed = False
            for source_mapping in edited_mappings:
                # Find the corresponding column by original header, ignoring case and whitespace
                for target_mapping in target_sheet.column_mappings:
                    if target_mapping.original_header.strip().lower() == source_mapping.original_header.strip().lower():
                        
                        # Propagate if the type is different, or if the target wasn't already manually edited
                        if (target_mapping.mapped_type != source_mapping.mapped_type or 
                            not target_mapping.is_user_edited):
                            
                            target_mapping.mapped_type = source_mapping.mapped_type
                            target_mapping.confidence = 1.0
                            target_mapping.is_user_edited = True
                            target_mapping.is_required = source_mapping.is_required
                            propagated_count += 1
                            sheet_was_changed = True

                            # Enforce uniqueness on the target sheet
                            if target_mapping.is_required:
                                for other_cm in target_sheet.column_mappings:
                                    if (other_cm.original_header.strip().lower() != target_mapping.original_header.strip().lower() and
                                        other_cm.mapped_type == target_mapping.mapped_type and
                                        other_cm.is_required):
                                        # Demote the conflicting column
                                        if len(other_cm.alternatives) > 1:
                                            second_best = other_cm.alternatives[1]
                                            other_cm.mapped_type = second_best['type']
                                            other_cm.confidence = second_best['confidence']
                                        else:
                                            other_cm.mapped_type = 'unknown'
                                            other_cm.confidence = 0.0
                                        other_cm.is_required = False
                        break 
            
            if sheet_was_changed:
                changed_sheets.add(target_sheet.sheet_name)

        # --- Refresh all affected tree views ---
        for sheet_name in changed_sheets:
            if sheet_name in self.sheet_treeviews:
                tree = self.sheet_treeviews[sheet_name]
                # Find the corresponding sheet mapping object
                sheet_map = next((s for s in self.file_mapping.sheets if s.sheet_name == sheet_name), None)
                if tree and sheet_map:
                    self._refresh_tree_view(tree, sheet_map)

        messagebox.showinfo("Propagation Complete", f"Applied {propagated_count} changes to {len(changed_sheets)} other sheets.")
        self._update_status("Propagated column mappings to other sheets.")

    def _edit_column_mapping(self, tree, sheet_mapping):
        """Edit column mapping when double-clicked"""
        selection = tree.selection()
        if not selection:
            return
            
        item = tree.item(selection[0])
        values = item['values']
        if not values:
            return
            
        column_name = values[0]
        current_type = values[1]
        
        # Find the column mapping object
        col_mapping = None
        for cm in sheet_mapping.column_mappings:
            if cm.original_header == column_name:
                col_mapping = cm
                break
        
        if not col_mapping:
            return
            
        # Create a dialog to edit the column type
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Column: {column_name}")
        dialog.geometry("500x600")  # Made taller
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Main frame with scrollbar
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        tk.Label(main_frame, text=f"Column: {column_name}", font=("Arial", 12, "bold")).pack(pady=10)
        ttk.Label(main_frame, text="Select new column type:").pack(pady=5)
        
        # Create a frame for the radio buttons with scrollbar
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Canvas and scrollbar for radio buttons
        canvas = tk.Canvas(radio_frame)
        scrollbar = ttk.Scrollbar(radio_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Column type options including "Ignore"
        from utils.config import ColumnType
        type_var = tk.StringVar(value=current_type)
        
        # Add "Ignore" option first
        tk.Radiobutton(scrollable_frame, text="Ignore (Skip this column)", 
                       variable=type_var, value="ignore", 
                       font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=2)
        
        ttk.Separator(scrollable_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        
        # Add all other column types
        for col_type in ColumnType:
            ttk.Radiobutton(scrollable_frame, text=col_type.value, 
                           variable=type_var, value=col_type.value).pack(anchor=tk.W, pady=1)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def save_changes():
            new_type = type_var.get()
            if new_type != current_type:
                # Check for uniqueness constraint on required columns
                if new_type != "ignore":
                    required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
                    if new_type in required_types:
                        # Find any existing column of this required type
                        existing_required_column = None
                        for cm in sheet_mapping.column_mappings:
                            if cm.mapped_type == new_type and cm.original_header != column_name and cm.is_required:
                                existing_required_column = cm
                                break
                        
                        if existing_required_column:
                            # Ask user if they want to replace the existing mapping
                            replace = messagebox.askyesno(
                                "Duplicate Required Column",
                                f"Column '{existing_required_column.original_header}' is already the required '{new_type}' column.\n\n"
                                f"Do you want to assign '{column_name}' as the new required column?\n\n"
                                f"This will change '{existing_required_column.original_header}' to its second-best guess.",
                                icon='warning'
                            )
                            
                            if replace:
                                # Demote the existing column to its second-best guess or unknown
                                # The first alternative (index 0) is the current mapping, so we try index 1.
                                if len(existing_required_column.alternatives) > 1:
                                    second_best = existing_required_column.alternatives[1]
                                    existing_required_column.mapped_type = second_best['type']
                                    existing_required_column.confidence = second_best['confidence']
                                else:
                                    # No second-best alternative, set to unknown
                                    existing_required_column.mapped_type = "unknown"
                                    existing_required_column.confidence = 0.0
                                
                                existing_required_column.is_required = False
                                existing_required_column.is_user_edited = True
                            else:
                                # User cancelled, don't make the change
                                dialog.destroy()
                                return
                
                # Update the current column mapping
                if new_type == "ignore":
                    # For ignored columns, we'll mark them specially
                    col_mapping.mapped_type = "ignore"
                    col_mapping.confidence = 1.0
                    col_mapping.is_required = False
                    col_mapping.is_user_edited = True
                else:
                    col_mapping.mapped_type = new_type
                    col_mapping.confidence = 1.0  # Manual override gets full confidence
                    # Update required status based on the new type
                    required_types = ["description", "quantity", "unit_price", "total_price", "unit", "code"]
                    col_mapping.is_required = new_type in required_types
                    col_mapping.is_user_edited = True
                
                # Enforce uniqueness for required columns after the change
                if col_mapping.is_required:
                    # Find any other columns of the same required type and demote them
                    for cm in sheet_mapping.column_mappings:
                        if (cm.original_header != col_mapping.original_header and 
                            cm.mapped_type == col_mapping.mapped_type and 
                            cm.is_required):
                            # Demote this column to its second-best guess or unknown
                            if len(cm.alternatives) > 1:
                                second_best = cm.alternatives[1]
                                cm.mapped_type = second_best['type']
                                cm.confidence = second_best['confidence']
                            else:
                                cm.mapped_type = "unknown"
                                cm.confidence = 0.0
                            cm.is_required = False
                            cm.is_user_edited = True
                
                # Refresh the entire tree view to ensure consistency
                self._refresh_tree_view(tree, sheet_mapping)
                
                self._update_status(f"Updated column '{column_name}' to type '{new_type}' (user edited)")
            
            dialog.destroy()
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save", command=save_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def _show_row_details(self, tree, sheet_mapping):
        """Show detailed row classifications when double-clicked"""
        selection = tree.selection()
        if not selection:
            return
            
        item = tree.item(selection[0])
        values = item['values']
        if not values:
            return
            
        row_type = values[0]
        
        # Create a dialog to show detailed row classifications
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Row Details: {row_type}")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Filter rows by type
        type_rows = [r for r in sheet_mapping.row_classifications if r.row_type == row_type]
        
        # Create a frame with scrollbar
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tree view for detailed rows
        columns = ["Row Index", "Type", "Confidence", "Completeness", "Section Title", "Validation Errors"]
        detail_tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        for col in columns:
            detail_tree.heading(col, text=col)
            detail_tree.column(col, width=120)
        
        for row_class in type_rows:
            detail_tree.insert('', tk.END, values=[
                row_class.row_index,
                row_class.row_type,
                f"{row_class.confidence:.1%}",
                f"{row_class.completeness_score:.1%}",
                row_class.section_title or "",
                len(row_class.validation_errors)
            ])
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=scrollbar.set)
        
        detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add close button
        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=10)

    def export_file(self):
        """Handle the file export process."""
        if not self.open_files:
            messagebox.showwarning("No File Open", "Please open a file before exporting.")
            return

        # Get the current active file path
        current_tab_widget = self.notebook.nametowidget(self.notebook.select())
        filepath = None
        for path, tab in self.open_files.items():
            if tab == current_tab_widget:
                filepath = path
                break

        if not filepath:
            messagebox.showerror("Error", "Could not determine the active file to export.")
            return
            
        filetypes = [
            ("Normalized Excel", "*.xlsx"),
            ("Summary Report Excel", "*.xlsx"),
            ("JSON Data", "*.json"),
            ("CSV Data", "*.csv"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.asksaveasfilename(
            title="Export Processed Data",
            defaultextension=".xlsx",
            filetypes=filetypes
        )

        if not filename:
            return  # User cancelled

        # Determine selected format
        selected_type = (filename.split('.')[-1]).lower()
        if "summary" in filename.lower():
             format_type = 'summary_excel'
        elif selected_type == 'xlsx':
            format_type = 'normalized_excel'
        elif selected_type == 'json':
            format_type = 'json'
        elif selected_type == 'csv':
            format_type = 'csv'
        else:
            format_type = 'normalized_excel' # Default

        try:
            self._update_status(f"Exporting to {filename}...")
            success = self.controller.export_file(filepath, Path(filename), format_type)
            if success:
                self._update_status(f"Successfully exported to {filename}")
                messagebox.showinfo("Export Successful", f"Data exported to:\n{filename}")
            else:
                self._update_status("Export failed.")
                messagebox.showerror("Export Failed", "Could not export the file. Check logs for details.")
        except Exception as e:
            self._update_status("Export error.")
            messagebox.showerror("Export Error", f"An error occurred during export:\n{e}")
            logger.error(f"GUI export failed: {e}", exc_info=True)

    def open_settings(self):
        """Open the settings dialog"""
        if SETTINGS_AVAILABLE:
            try:
                # Get current settings (you can pass actual settings here)
                current_settings = {}
                show_settings_dialog(self.root, current_settings, self._on_settings_save)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open settings: {e}")
        else:
            messagebox.showinfo("Not Available", "Settings dialog is not available.")

    def _on_settings_save(self, new_settings):
        """Callback when settings are saved"""
        self._update_status("Settings saved successfully!")
        # Here you can apply the new settings to your application
        print("New settings:", new_settings)

    def _refresh_tree_view(self, tree, sheet_mapping):
        """Refresh the entire tree view to ensure consistency"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Re-populate with current data
        for col_mapping in sheet_mapping.column_mappings:
            # Determine tags based on current state
            tags = []
            if col_mapping.is_required:
                tags.append('required')
            if col_mapping.is_user_edited:
                tags.append('edited')
            
            # Determine required text
            if col_mapping.mapped_type == "ignore":
                required_text = "Ignore"
            elif col_mapping.is_required:
                required_text = "Yes"
            else:
                required_text = "No"
            
            # Determine actions text
            if col_mapping.is_user_edited:
                actions_text = "Edited"
            else:
                actions_text = "Auto-detected"
            
            tree.insert('', 'end', values=(
                col_mapping.original_header,
                col_mapping.mapped_type,
                f"{col_mapping.confidence:.1%}",
                required_text,
                actions_text
            ), tags=tuple(tags))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    # This block is for testing the UI component independently.
    # A mock controller would be needed to test full functionality.
    class MockController:
        def export_file(self, *args, **kwargs):
            print("MockController: export_file called")
            return True
            
    app = MainWindow(controller=MockController())
    app.run() 