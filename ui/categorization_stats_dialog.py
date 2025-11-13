"""
Categorization Statistics Dialog for BOQ Tools
Shows detailed categorization statistics and coverage reports
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from typing import Dict, List, Any, Optional
import logging
import numpy as np

# Optional matplotlib import
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

logger = logging.getLogger(__name__)


class CategorizationStatsDialog:
    def __init__(self, parent, dataframe, categorization_result=None):
        """
        Initialize the categorization statistics dialog
        
        Args:
            parent: Parent window
            dataframe: DataFrame with categorization data
            categorization_result: Result from categorization process
        """
        self.parent = parent
        self.dataframe = dataframe
        self.categorization_result = categorization_result
        
        # Dialog state
        self.dialog = None
        self.notebook = None
        
        # Create and show dialog
        self._create_dialog()
    
    def _create_dialog(self):
        """Create the main dialog window"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Categorization Statistics")
        self.dialog.geometry("900x700")
        self.dialog.minsize(800, 600)
        
        # Make dialog modal
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog on parent
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"900x700+{x}+{y}")
        
        # Create main content
        self._create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.dialog.destroy)
    
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
        title_label = ttk.Label(main_frame, text="Categorization Statistics & Coverage Report", 
                               font=("TkDefaultFont", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=tk.NSEW, pady=(0, 10))
        
        # Create tabs
        self._create_summary_tab()
        self._create_coverage_tab()
        self._create_category_breakdown_tab()
        
        # Only create charts tab if matplotlib is available
        if MATPLOTLIB_AVAILABLE:
            self._create_charts_tab()
        else:
            self._create_no_charts_tab()
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, sticky=tk.EW, pady=(10, 0))
        
        export_button = ttk.Button(button_frame, text="Export Report", 
                                  command=self._export_report)
        export_button.pack(side=tk.LEFT)
        
        close_button = ttk.Button(button_frame, text="Close", 
                                 command=self.dialog.destroy)
        close_button.pack(side=tk.RIGHT)
        
        # Configure button frame
        button_frame.grid_columnconfigure(0, weight=1)
    
    def _create_summary_tab(self):
        """Create the summary tab"""
        summary_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(summary_frame, text="Summary")
        
        # Configure grid
        summary_frame.grid_columnconfigure(0, weight=1)
        
        # Summary statistics
        stats = self._calculate_summary_stats()
        
        # Create summary display
        row = 0
        
        # Overall statistics
        overall_frame = ttk.LabelFrame(summary_frame, text="Overall Statistics", padding="10")
        overall_frame.grid(row=row, column=0, sticky=tk.EW, pady=(0, 10))
        row += 1
        
        ttk.Label(overall_frame, text=f"Total Rows: {stats['total_rows']}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Label(overall_frame, text=f"Categorized Rows: {stats['categorized_rows']}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Label(overall_frame, text=f"Coverage Rate: {stats['coverage_rate']:.1%}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=2, sticky=tk.W)
        
        # Category statistics
        category_frame = ttk.LabelFrame(summary_frame, text="Category Statistics", padding="10")
        category_frame.grid(row=row, column=0, sticky=tk.EW, pady=(0, 10))
        row += 1
        
        ttk.Label(category_frame, text=f"Unique Categories: {stats['unique_categories']}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Label(category_frame, text=f"Most Common Category: {stats['most_common_category']}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Label(category_frame, text=f"Average Category Size: {stats['avg_category_size']:.1f}", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=0, column=2, sticky=tk.W)
        
        # Process statistics (if available)
        if self.categorization_result:
            process_frame = ttk.LabelFrame(summary_frame, text="Process Statistics", padding="10")
            process_frame.grid(row=row, column=0, sticky=tk.EW, pady=(0, 10))
            row += 1
            
            all_stats = self.categorization_result.get('all_stats', {})
            
            # Auto-categorization stats
            auto_stats = all_stats.get('auto_stats', {})
            if auto_stats:
                ttk.Label(process_frame, text="Auto-Categorization:", 
                         font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
                ttk.Label(process_frame, text=f"Matched: {auto_stats.get('matched_rows', 0)}").grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
                ttk.Label(process_frame, text=f"Rate: {auto_stats.get('match_rate', 0):.1%}").grid(row=0, column=2, sticky=tk.W)
            
            # Manual categorization stats
            apply_stats = all_stats.get('apply_stats', {})
            if apply_stats:
                ttk.Label(process_frame, text="Manual Categorization:", 
                         font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
                ttk.Label(process_frame, text=f"Updated: {apply_stats.get('rows_updated', 0)}").grid(row=1, column=1, sticky=tk.W, padx=(0, 20))
                ttk.Label(process_frame, text=f"Final Coverage: {apply_stats.get('coverage_rate', 0):.1%}").grid(row=1, column=2, sticky=tk.W)
        
        # Configure grid weights
        overall_frame.grid_columnconfigure(2, weight=1)
        category_frame.grid_columnconfigure(2, weight=1)
        if self.categorization_result:
            process_frame.grid_columnconfigure(2, weight=1)
    
    def _create_coverage_tab(self):
        """Create the coverage tab"""
        coverage_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(coverage_frame, text="Coverage")
        
        # Configure grid
        coverage_frame.grid_columnconfigure(0, weight=1)
        coverage_frame.grid_rowconfigure(1, weight=1)
        
        # Title
        ttk.Label(coverage_frame, text="Coverage by Source Sheet", 
                 font=("TkDefaultFont", 12, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Create treeview for coverage details
        columns = ('Sheet', 'Total_Rows', 'Categorized', 'Uncategorized', 'Coverage_Rate')
        tree = ttk.Treeview(coverage_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        tree.heading('Sheet', text='Source Sheet')
        tree.heading('Total_Rows', text='Total Rows')
        tree.heading('Categorized', text='Categorized')
        tree.heading('Uncategorized', text='Uncategorized')
        tree.heading('Coverage_Rate', text='Coverage Rate')
        
        tree.column('Sheet', width=200, minwidth=150)
        tree.column('Total_Rows', width=100, minwidth=80)
        tree.column('Categorized', width=100, minwidth=80)
        tree.column('Uncategorized', width=100, minwidth=80)
        tree.column('Coverage_Rate', width=120, minwidth=100)
        
        # Scrollbars
        vsb = ttk.Scrollbar(coverage_frame, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(coverage_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        tree.grid(row=1, column=0, sticky=tk.NSEW)
        vsb.grid(row=1, column=1, sticky=tk.NS)
        hsb.grid(row=2, column=0, sticky=tk.EW)
        
        # Load coverage data
        self._load_coverage_data(tree)
    
    def _create_category_breakdown_tab(self):
        """Create the category breakdown tab"""
        breakdown_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(breakdown_frame, text="Category Breakdown")
        
        # Configure grid
        breakdown_frame.grid_columnconfigure(0, weight=1)
        breakdown_frame.grid_rowconfigure(1, weight=1)
        
        # Title
        ttk.Label(breakdown_frame, text="Category Distribution", 
                 font=("TkDefaultFont", 12, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Create treeview for category breakdown
        columns = ('Category', 'Count', 'Percentage', 'Avg_Description_Length')
        tree = ttk.Treeview(breakdown_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        tree.heading('Category', text='Category')
        tree.heading('Count', text='Count')
        tree.heading('Percentage', text='Percentage')
        tree.heading('Avg_Description_Length', text='Avg Description Length')
        
        tree.column('Category', width=250, minwidth=200)
        tree.column('Count', width=100, minwidth=80)
        tree.column('Percentage', width=100, minwidth=80)
        tree.column('Avg_Description_Length', width=150, minwidth=120)
        
        # Scrollbars
        vsb = ttk.Scrollbar(breakdown_frame, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(breakdown_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        tree.grid(row=1, column=0, sticky=tk.NSEW)
        vsb.grid(row=1, column=1, sticky=tk.NS)
        hsb.grid(row=2, column=0, sticky=tk.EW)
        
        # Load category breakdown data
        self._load_category_breakdown_data(tree)
    
    def _create_charts_tab(self):
        """Create the charts tab with matplotlib"""
        charts_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(charts_frame, text="Charts")
        
        # Configure grid
        charts_frame.grid_columnconfigure(0, weight=1)
        charts_frame.grid_columnconfigure(1, weight=1)
        charts_frame.grid_rowconfigure(1, weight=1)
        
        # Title
        ttk.Label(charts_frame, text="Visualizations", 
                 font=("TkDefaultFont", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Create charts
        self._create_coverage_chart(charts_frame)
        self._create_category_chart(charts_frame)
    
    def _create_no_charts_tab(self):
        """Create a tab explaining that charts are not available"""
        no_charts_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(no_charts_frame, text="Charts")
        
        # Configure grid
        no_charts_frame.grid_columnconfigure(0, weight=1)
        no_charts_frame.grid_rowconfigure(0, weight=1)
        
        # Message
        message = """
        Charts are not available because matplotlib is not installed.
        
        To enable charts, install matplotlib:
        pip install matplotlib
        
        Charts provide visual representations of:
        - Coverage distribution (pie chart)
        - Top categories (bar chart)
        - Category trends over time
        """
        
        message_label = ttk.Label(no_charts_frame, text=message, 
                                 wraplength=400, justify=tk.LEFT)
        message_label.grid(row=0, column=0, sticky=tk.NSEW, pady=20)
    
    def _create_coverage_chart(self, parent):
        """Create coverage pie chart"""
        if not MATPLOTLIB_AVAILABLE:
            return
            
        chart_frame = ttk.LabelFrame(parent, text="Coverage Distribution", padding="10")
        chart_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, 5))
        
        # Create matplotlib figure
        fig, ax = plt.subplots(figsize=(6, 4))
        
        # Get coverage data
        stats = self._calculate_summary_stats()
        categorized = stats['categorized_rows']
        uncategorized = stats['total_rows'] - stats['categorized_rows']
        
        # Create pie chart
        labels = ['Categorized', 'Uncategorized']
        sizes = [categorized, uncategorized]
        colors = ['#4CAF50', '#F44336']
        
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _create_category_chart(self, parent):
        """Create category bar chart"""
        if not MATPLOTLIB_AVAILABLE:
            return
            
        chart_frame = ttk.LabelFrame(parent, text="Top Categories", padding="10")
        chart_frame.grid(row=1, column=1, sticky=tk.NSEW, padx=(5, 0))
        
        # Create matplotlib figure
        fig, ax = plt.subplots(figsize=(6, 4))
        
        # Get top categories
        if 'Category' in self.dataframe.columns:
            category_counts = self.dataframe['Category'].value_counts().head(10)
            
            # Create bar chart
            categories = category_counts.index.tolist()
            counts = category_counts.values.tolist()
            
            y_pos = np.arange(len(categories))
            ax.barh(y_pos, counts)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(categories)
            ax.set_xlabel('Count')
            ax.set_title('Top 10 Categories')
            
            # Rotate labels if needed
            plt.setp(ax.get_yticklabels(), rotation=0, ha='right')
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _calculate_summary_stats(self) -> Dict[str, Any]:
        """Calculate summary statistics"""
        total_rows = len(self.dataframe)
        
        if 'Category' in self.dataframe.columns:
            categorized_mask = self.dataframe['Category'].notna() & (self.dataframe['Category'] != '')
            categorized_rows = categorized_mask.sum()
            coverage_rate = categorized_rows / total_rows if total_rows > 0 else 0
            
            # Category statistics
            category_counts = self.dataframe['Category'].value_counts()
            unique_categories = len(category_counts)
            most_common_category = category_counts.index[0] if len(category_counts) > 0 else 'None'
            avg_category_size = category_counts.mean() if len(category_counts) > 0 else 0
            
            # Average description length
            if 'Description' in self.dataframe.columns:
                avg_desc_length = self.dataframe['Description'].astype(str).str.len().mean()
            else:
                avg_desc_length = 0
        else:
            categorized_rows = 0
            coverage_rate = 0
            unique_categories = 0
            most_common_category = 'None'
            avg_category_size = 0
            avg_desc_length = 0
        
        return {
            'total_rows': total_rows,
            'categorized_rows': categorized_rows,
            'coverage_rate': coverage_rate,
            'unique_categories': unique_categories,
            'most_common_category': most_common_category,
            'avg_category_size': avg_category_size,
            'avg_description_length': avg_desc_length
        }
    
    def _load_coverage_data(self, tree):
        """Load coverage data into treeview"""
        if 'Source_Sheet' not in self.dataframe.columns:
            return
        
        # Group by source sheet
        sheet_stats = []
        for sheet in self.dataframe['Source_Sheet'].unique():
            sheet_data = self.dataframe[self.dataframe['Source_Sheet'] == sheet]
            total_rows = len(sheet_data)
            
            if 'Category' in sheet_data.columns:
                categorized_mask = sheet_data['Category'].notna() & (sheet_data['Category'] != '')
                categorized = categorized_mask.sum()
                uncategorized = total_rows - categorized
                coverage_rate = categorized / total_rows if total_rows > 0 else 0
            else:
                categorized = 0
                uncategorized = total_rows
                coverage_rate = 0
            
            sheet_stats.append({
                'sheet': sheet,
                'total': total_rows,
                'categorized': categorized,
                'uncategorized': uncategorized,
                'coverage_rate': coverage_rate
            })
        
        # Sort by coverage rate
        sheet_stats.sort(key=lambda x: x['coverage_rate'], reverse=True)
        
        # Insert into treeview
        for stats in sheet_stats:
            tree.insert('', 'end', values=(
                stats['sheet'],
                stats['total'],
                stats['categorized'],
                stats['uncategorized'],
                f"{stats['coverage_rate']:.1%}"
            ))
    
    def _load_category_breakdown_data(self, tree):
        """Load category breakdown data into treeview"""
        if 'Category' not in self.dataframe.columns:
            return
        
        # Get category statistics
        category_counts = self.dataframe['Category'].value_counts()
        total_rows = len(self.dataframe)
        
        # Calculate average description length per category
        avg_lengths = {}
        if 'Description' in self.dataframe.columns:
            for category in category_counts.index:
                category_data = self.dataframe[self.dataframe['Category'] == category]
                avg_lengths[category] = category_data['Description'].astype(str).str.len().mean()
        
        # Insert into treeview
        for category, count in category_counts.items():
            percentage = count / total_rows if total_rows > 0 else 0
            avg_length = avg_lengths.get(category, 0)
            
            tree.insert('', 'end', values=(
                category,
                count,
                f"{percentage:.1%}",
                f"{avg_length:.1f}"
            ))
    
    def _export_report(self):
        """Export the statistics report"""
        from tkinter import filedialog
        
        file_path = filedialog.asksaveasfilename(
            title="Export Statistics Report",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Create a comprehensive report
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Summary sheet
                    summary_data = self._create_summary_sheet()
                    summary_data.to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Coverage sheet
                    coverage_data = self._create_coverage_sheet()
                    coverage_data.to_excel(writer, sheet_name='Coverage', index=False)
                    
                    # Category breakdown sheet
                    breakdown_data = self._create_breakdown_sheet()
                    breakdown_data.to_excel(writer, sheet_name='Category_Breakdown', index=False)
                    
                    # Raw data sheet
                    self.dataframe.to_excel(writer, sheet_name='Raw_Data', index=False)
                
                messagebox.showinfo("Success", f"Report exported to: {file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export report: {str(e)}")
    
    def _create_summary_sheet(self) -> pd.DataFrame:
        """Create summary sheet data"""
        stats = self._calculate_summary_stats()
        
        data = {
            'Metric': [
                'Total Rows',
                'Categorized Rows',
                'Uncategorized Rows',
                'Coverage Rate',
                'Unique Categories',
                'Most Common Category',
                'Average Category Size',
                'Average Description Length'
            ],
            'Value': [
                stats['total_rows'],
                stats['categorized_rows'],
                stats['total_rows'] - stats['categorized_rows'],
                f"{stats['coverage_rate']:.1%}",
                stats['unique_categories'],
                stats['most_common_category'],
                f"{stats['avg_category_size']:.1f}",
                f"{stats['avg_description_length']:.1f}"
            ]
        }
        
        return pd.DataFrame(data)
    
    def _create_coverage_sheet(self) -> pd.DataFrame:
        """Create coverage sheet data"""
        if 'Source_Sheet' not in self.dataframe.columns:
            return pd.DataFrame()
        
        sheet_stats = []
        for sheet in self.dataframe['Source_Sheet'].unique():
            sheet_data = self.dataframe[self.dataframe['Source_Sheet'] == sheet]
            total_rows = len(sheet_data)
            
            if 'Category' in sheet_data.columns:
                categorized_mask = sheet_data['Category'].notna() & (sheet_data['Category'] != '')
                categorized = categorized_mask.sum()
                uncategorized = total_rows - categorized
                coverage_rate = categorized / total_rows if total_rows > 0 else 0
            else:
                categorized = 0
                uncategorized = total_rows
                coverage_rate = 0
            
            sheet_stats.append({
                'Source_Sheet': sheet,
                'Total_Rows': total_rows,
                'Categorized': categorized,
                'Uncategorized': uncategorized,
                'Coverage_Rate': coverage_rate
            })
        
        return pd.DataFrame(sheet_stats)
    
    def _create_breakdown_sheet(self) -> pd.DataFrame:
        """Create category breakdown sheet data"""
        if 'Category' not in self.dataframe.columns:
            return pd.DataFrame()
        
        category_counts = self.dataframe['Category'].value_counts()
        total_rows = len(self.dataframe)
        
        # Calculate average description length per category
        avg_lengths = {}
        if 'Description' in self.dataframe.columns:
            for category in category_counts.index:
                category_data = self.dataframe[self.dataframe['Category'] == category]
                avg_lengths[category] = category_data['Description'].astype(str).str.len().mean()
        
        data = {
            'Category': category_counts.index.tolist(),
            'Count': category_counts.values.tolist(),
            'Percentage': [(count / total_rows) * 100 for count in category_counts.values],
            'Average_Description_Length': [avg_lengths.get(cat, 0) for cat in category_counts.index]
        }
        
        return pd.DataFrame(data)


def show_categorization_stats_dialog(parent, dataframe, categorization_result=None):
    """
    Show the categorization statistics dialog
    
    Args:
        parent: Parent window
        dataframe: DataFrame with categorization data
        categorization_result: Result from categorization process
    
    Returns:
        CategorizationStatsDialog instance
    """
    dialog = CategorizationStatsDialog(parent, dataframe, categorization_result)
    return dialog 