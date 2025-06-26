#!/usr/bin/env python3
"""
UI Categorization Integration Demo
Demonstrates the complete categorization workflow with UI integration
"""

import sys
import os
from pathlib import Path
import pandas as pd
import logging

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Core components
from core.auto_categorizer import CategoryDictionary
from core.manual_categorizer import execute_row_categorization
from ui.categorization_dialog import show_categorization_dialog
from ui.category_review_dialog import show_category_review_dialog
from ui.categorization_stats_dialog import show_categorization_stats_dialog

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class MockController:
    """Mock controller for demo purposes"""
    
    def __init__(self):
        self.current_files = {}
    
    def get_current_dataframe(self):
        """Get current DataFrame for demo"""
        return self.demo_dataframe
    
    def update_progress(self, percent, message):
        """Update progress for demo"""
        print(f"Progress: {percent}% - {message}")


class MockFileMapping:
    """Mock file mapping for demo purposes"""
    
    def __init__(self, dataframe):
        self.dataframe = dataframe
        self.categorized_dataframe = None
        self.categorization_result = None
    
    def get_processed_dataframe(self):
        """Get processed DataFrame"""
        return self.dataframe


def create_demo_data():
    """Create demo data for testing"""
    data = {
        'Description': [
            'Concrete foundation work',
            'Steel reinforcement bars',
            'Electrical wiring installation',
            'Plumbing pipe installation',
            'Roofing materials',
            'Window installation',
            'Door frames and hardware',
            'Flooring materials',
            'Wall finishing',
            'HVAC system installation',
            'Unknown construction item',
            'Miscellaneous materials',
            'Labor costs',
            'Equipment rental',
            'Site preparation work'
        ],
        'Quantity': [100, 500, 200, 150, 80, 20, 15, 200, 300, 1, 50, 100, 1000, 30, 200],
        'Unit': ['m³', 'kg', 'm', 'm', 'm²', 'units', 'units', 'm²', 'm²', 'system', 'units', 'kg', 'hours', 'days', 'm²'],
        'Unit_Price': [150, 2.5, 15, 25, 45, 200, 150, 35, 20, 5000, 10, 5, 25, 100, 30],
        'Total_Price': [15000, 1250, 3000, 3750, 3600, 4000, 2250, 7000, 6000, 5000, 500, 500, 25000, 3000, 6000],
        'Source_Sheet': ['BOQ_Sheet_1'] * 15
    }
    
    return pd.DataFrame(data)


def demo_categorization_workflow():
    """Demonstrate the complete categorization workflow"""
    print("=== UI Categorization Integration Demo ===\n")
    
    # Create demo data
    print("1. Creating demo data...")
    demo_df = create_demo_data()
    print(f"   Created DataFrame with {len(demo_df)} rows")
    print(f"   Columns: {list(demo_df.columns)}")
    print()
    
    # Create mock controller and file mapping
    controller = MockController()
    controller.demo_dataframe = demo_df
    file_mapping = MockFileMapping(demo_df)
    
    # Demo 1: Execute categorization workflow
    print("2. Executing categorization workflow...")
    try:
        result = execute_row_categorization(
            mapped_df=demo_df,
            progress_callback=controller.update_progress
        )
        
        if result['error']:
            print(f"   Error: {result['error']}")
        else:
            print("   Categorization completed successfully!")
            print(f"   Final DataFrame shape: {result['final_dataframe'].shape}")
            print(f"   Categories found: {result['final_dataframe']['Category'].nunique()}")
            
            # Update file mapping
            file_mapping.categorized_dataframe = result['final_dataframe']
            file_mapping.categorization_result = result
            
    except Exception as e:
        print(f"   Error during categorization: {e}")
        return
    
    print()
    
    # Demo 2: Show categorization statistics
    print("3. Showing categorization statistics...")
    try:
        # This would normally be called from the UI
        # For demo, we'll just print the statistics
        stats = result.get('all_stats', {})
        
        print("   Auto-categorization stats:")
        auto_stats = stats.get('auto_stats', {})
        if auto_stats:
            print(f"     Total rows: {auto_stats.get('total_rows', 0)}")
            print(f"     Matched rows: {auto_stats.get('matched_rows', 0)}")
            print(f"     Match rate: {auto_stats.get('match_rate', 0):.1%}")
        
        print("   Manual categorization stats:")
        apply_stats = stats.get('apply_stats', {})
        if apply_stats:
            print(f"     Rows updated: {apply_stats.get('rows_updated', 0)}")
            print(f"     Final coverage: {apply_stats.get('coverage_rate', 0):.1%}")
        
        print("   Dictionary updates:")
        update_result = stats.get('update_result', {})
        if update_result:
            print(f"     New mappings added: {update_result.get('total_added', 0)}")
            print(f"     Conflicts found: {update_result.get('total_conflicts', 0)}")
            
    except Exception as e:
        print(f"   Error showing statistics: {e}")
    
    print()
    
    # Demo 3: Show category distribution
    print("4. Category distribution:")
    if 'Category' in result['final_dataframe'].columns:
        category_counts = result['final_dataframe']['Category'].value_counts()
        for category, count in category_counts.items():
            print(f"   {category}: {count} rows")
    
    print()
    
    # Demo 4: Export categorized data
    print("5. Exporting categorized data...")
    try:
        output_file = "examples/ui_categorization_demo_output.csv"
        result['final_dataframe'].to_csv(output_file, index=False)
        print(f"   Exported to: {output_file}")
    except Exception as e:
        print(f"   Error exporting: {e}")
    
    print()
    print("=== Demo completed successfully! ===")
    print("\nNote: To test the actual UI dialogs, run the main application")
    print("and use the categorization features through the GUI.")


def demo_ui_dialogs():
    """Demonstrate the UI dialogs (requires GUI)"""
    print("=== UI Dialog Demo ===")
    print("This demo requires a GUI environment.")
    print("To test the UI dialogs:")
    print("1. Run the main application: python main.py --gui")
    print("2. Load an Excel file")
    print("3. Complete the column mapping and row review steps")
    print("4. Click 'Confirm Row Review' to start categorization")
    print("5. Use the categorization dialogs to:")
    print("   - Review and modify categories")
    print("   - View statistics and coverage reports")
    print("   - Export categorized data")


if __name__ == "__main__":
    print("BOQ Tools - UI Categorization Integration Demo")
    print("=" * 50)
    
    # Check if we're in a GUI environment
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Hide the window
        GUI_AVAILABLE = True
    except:
        GUI_AVAILABLE = False
        print("GUI not available, running console demo only")
    
    # Run the workflow demo
    demo_categorization_workflow()
    
    if GUI_AVAILABLE:
        print("\n" + "=" * 50)
        demo_ui_dialogs()
    
    print("\nDemo completed!") 