#!/usr/bin/env python3
"""
Test script to check the comparison workflow
"""

import sys
import os
import logging

# Add the project root to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    handlers=[
        logging.FileHandler('comparison_workflow_test.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

def test_comparison_workflow():
    """Test the comparison workflow to see which dialogs are being called"""
    try:
        from ui.main_window import MainWindow
        from core.controller import Controller
        
        # Create controller and main window
        controller = Controller()
        main_window = MainWindow(controller)
        
        logger.info("Testing comparison workflow...")
        
        # Simulate the comparison workflow
        # This will help us see which dialogs are being called
        
        # Check if COMPARISON_ROW_REVIEW_AVAILABLE is True
        from ui.main_window import COMPARISON_ROW_REVIEW_AVAILABLE
        logger.info(f"COMPARISON_ROW_REVIEW_AVAILABLE: {COMPARISON_ROW_REVIEW_AVAILABLE}")
        
        # Check if the comparison dialog can be imported
        try:
            from ui.comparison_row_review_dialog import show_comparison_row_review
            logger.info("show_comparison_row_review imported successfully")
        except ImportError as e:
            logger.error(f"Failed to import show_comparison_row_review: {e}")
        
        logger.info("Test completed")
        
    except Exception as e:
        logger.error(f"Error in test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_comparison_workflow() 