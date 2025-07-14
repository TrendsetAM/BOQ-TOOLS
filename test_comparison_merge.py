#!/usr/bin/env python3
"""
Test script to verify the comparison/merge functionality works correctly.
This test simulates loading two BOQs and merging them for comparison.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk
import pandas as pd

# Add the project root to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.main_window import MainWindow

def test_comparison_merge():
    """Test the comparison/merge functionality"""
    print("Testing comparison/merge functionality...")
    
    # Create a mock controller
    class MockController:
        def __init__(self):
            self.current_files = {}
    
    controller = MockController()
    
    # Create the main window
    root = tk.Tk()
    root.withdraw()  # Hide the window
    
    main_window = MainWindow(controller, root)
    
    # Simulate the first BOQ being loaded
    print("Simulating first BOQ load...")
    
    # Set up offer info for first BOQ
    main_window.current_offer_info = {
        'supplier_name': 'Test Supplier 1',
        'project_name': 'Test Project 1',
        'date': '2025-01-15'
    }
    main_window.current_offer_name = 'Test Supplier 1'
    
    # Create a mock file mapping for first BOQ
    class MockFileMapping:
        def __init__(self):
            self.sheets = []
            self.categorized_dataframe = None
            self.tab = None
    
    file_mapping_1 = MockFileMapping()
    
    # Create a mock DataFrame with the first BOQ data
    df_1 = pd.DataFrame({
        'description': ['Item 1', 'Item 2', 'Item 3'],
        'quantity': [10, 5, 2],
        'unit_price': [100, 200, 150],
        'total_price': [1000, 1000, 300],
        'unit': ['pcs', 'm2', 'm'],
        'code': ['001', '002', '003'],
        'category': ['Category A', 'Category B', 'Category A']
    })
    
    # Set the categorized dataframe
    setattr(file_mapping_1, 'categorized_dataframe', df_1)
    
    # Add the first file to the controller's current files
    file_key_1 = "/test/path/boq_111.xlsx"
    controller.current_files[file_key_1] = {
        'file_mapping': file_mapping_1,
        'offer_info': {
            'supplier_name': 'Test Supplier 1',
            'project_name': 'Test Project 1', 
            'date': '2025-01-15'
        }
    }
    
    # Create a mock tab
    tab = ttk.Frame(root)
    file_mapping_1.tab = tab
    
    # Test the summary data collection for first BOQ
    print("Testing _collect_summary_data for first BOQ...")
    summary_data_1 = main_window._collect_summary_data()
    print(f"Collected summary data for first BOQ: {summary_data_1}")
    
    # Now simulate the second BOQ being loaded and merged
    print("Simulating second BOQ load and merge...")
    
    # Set up offer info for second BOQ
    main_window.current_offer_info = {
        'supplier_name': 'Test Supplier 2',
        'project_name': 'Test Project 2',
        'date': '2025-01-16'
    }
    main_window.current_offer_name = 'Test Supplier 2'
    
    # Create a mock file mapping for second BOQ
    file_mapping_2 = MockFileMapping()
    
    # Create a mock DataFrame with the second BOQ data (same structure, different prices)
    df_2 = pd.DataFrame({
        'description': ['Item 1', 'Item 2', 'Item 3'],
        'quantity': [10, 5, 2],
        'unit_price': [120, 220, 170],
        'total_price': [1200, 1100, 340],
        'unit': ['pcs', 'm2', 'm'],
        'code': ['001', '002', '003'],
        'category': ['Category A', 'Category B', 'Category A']
    })
    
    # Set the categorized dataframe
    setattr(file_mapping_2, 'categorized_dataframe', df_2)
    
    # Simulate the merge by creating a comparison DataFrame
    merged_df = pd.DataFrame({
        'description': ['Item 1', 'Item 2', 'Item 3'],
        'quantity': [10, 5, 2],
        'unit_price': [100, 200, 150],
        'total_price[Test Supplier 1]': [1000, 1000, 300],
        'total_price[Test Supplier 2]': [1200, 1100, 340],
        'unit': ['pcs', 'm2', 'm'],
        'code': ['001', '002', '003'],
        'category': ['Category A', 'Category B', 'Category A']
    })
    
    # Update the first file mapping with the merged DataFrame
    file_mapping_1.categorized_dataframe = merged_df
    controller.current_files[file_key_1]['categorized_dataframe'] = merged_df
    
    # Test the summary data collection after merge
    print("Testing _collect_summary_data after merge...")
    summary_data_2 = main_window._collect_summary_data()
    print(f"Collected summary data after merge: {summary_data_2}")
    
    # Verify that both offers are in the summary data
    if len(summary_data_2) == 2:
        print("✅ SUCCESS: Both offers are in the summary data")
        for i, (supplier, project, date, total_price) in enumerate(summary_data_2):
            print(f"  Offer {i+1}: {supplier}, {project}, {date}, {total_price}")
    else:
        print(f"❌ FAILURE: Expected 2 offers, got {len(summary_data_2)}")
    
    # Test the summary grid creation
    print("Testing _create_new_summary_grid after merge...")
    
    # Create a mock parent frame
    parent_frame = ttk.Frame(root)
    
    # Test the summary grid creation
    try:
        main_window._create_new_summary_grid(parent_frame, tab)
        print("✅ SUCCESS: Summary grid created without errors after merge")
    except Exception as e:
        print(f"❌ FAILURE: Error creating summary grid after merge: {e}")
        import traceback
        traceback.print_exc()
    
    root.destroy()
    print("Test completed.")

if __name__ == "__main__":
    test_comparison_merge() 