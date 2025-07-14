#!/usr/bin/env python3
"""
Test script to verify the summary grid refresh fix works correctly.
This test simulates the scenario where the first BOQ is loaded and the summary grid
should display the supplier/project/date/total price information.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk
import pandas as pd

# Add the project root to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.main_window import MainWindow
from core.boq_processor import BOQProcessor

def test_summary_grid_refresh():
    """Test the summary grid refresh functionality"""
    print("Testing summary grid refresh fix...")
    
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
    
    # Set up offer info (this would normally be done by the offer info dialog)
    main_window.current_offer_info = {
        'supplier_name': 'Test Supplier',
        'project_name': 'Test Project',
        'date': '2025-01-15'
    }
    main_window.current_offer_name = 'Test Supplier'
    
    # Create a mock file mapping
    class MockFileMapping:
        def __init__(self):
            self.sheets = []
            self.categorized_dataframe = None
    
    file_mapping = MockFileMapping()
    
    # Create a mock DataFrame with the first BOQ data
    df = pd.DataFrame({
        'description': ['Item 1', 'Item 2', 'Item 3'],
        'quantity': [10, 5, 2],
        'unit_price': [100, 200, 150],
        'total_price': [1000, 1000, 300],
        'unit': ['pcs', 'm2', 'm'],
        'code': ['001', '002', '003'],
        'category': ['Category A', 'Category B', 'Category A']
    })
    
    # Set the categorized dataframe
    setattr(file_mapping, 'categorized_dataframe', df)
    
    # Add the file to the controller's current files
    file_key = "/test/path/boq_111.xlsx"
    controller.current_files[file_key] = {
        'file_mapping': file_mapping,
        'offer_info': {
            'supplier_name': 'Test Supplier',
            'project_name': 'Test Project', 
            'date': '2025-01-15'
        }
    }
    
    # Test the summary data collection
    print("Testing _collect_summary_data...")
    summary_data = main_window._collect_summary_data()
    print(f"Collected summary data: {summary_data}")
    
    # Verify the summary data contains the expected information
    if summary_data:
        supplier, project_name, date, total_price = summary_data[0]
        print(f"Supplier: {supplier}")
        print(f"Project: {project_name}")
        print(f"Date: {date}")
        print(f"Total Price: {total_price}")
        
        # Check that we don't have "Unknown" values
        if supplier != 'Unknown' and project_name != 'Unknown' and date != 'Unknown':
            print("✅ SUCCESS: Summary data contains proper offer information")
        else:
            print("❌ FAILURE: Summary data contains 'Unknown' values")
    else:
        print("❌ FAILURE: No summary data collected")
    
    # Test the summary grid creation
    print("Testing _create_new_summary_grid...")
    
    # Create a mock tab and parent frame
    tab = ttk.Frame(root)
    parent_frame = ttk.Frame(root)
    
    # Test the summary grid creation
    try:
        main_window._create_new_summary_grid(parent_frame, tab)
        print("✅ SUCCESS: Summary grid created without errors")
    except Exception as e:
        print(f"❌ FAILURE: Error creating summary grid: {e}")
        import traceback
        traceback.print_exc()
    
    root.destroy()
    print("Test completed.")

if __name__ == "__main__":
    test_summary_grid_refresh() 