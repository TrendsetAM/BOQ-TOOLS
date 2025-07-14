#!/usr/bin/env python3
"""
Test script for summary grid refresh functionality
"""

import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_summary_refresh():
    """Test the summary grid refresh functionality"""
    print("Testing summary grid refresh functionality...")
    
    # Import the main window
    from ui.main_window import MainWindow
    from main import BOQApplicationController
    
    # Create controller
    controller = BOQApplicationController()
    
    # Create main window
    main_window = MainWindow(controller)
    
    # Test the centralized refresh method
    print("Testing _refresh_summary_grid_centralized...")
    main_window._refresh_summary_grid_centralized()
    
    # Test the clear files method
    print("Testing _clear_all_files...")
    main_window._clear_all_files()
    
    print("Test completed successfully!")

if __name__ == "__main__":
    test_summary_refresh() 