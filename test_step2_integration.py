#!/usr/bin/env python3
"""
Test script for Step 2: Integration of new row validity logic
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_main_window_integration():
    """Test that the MainWindow has the new row validity variables and logic"""
    print("=== Testing MainWindow Integration ===")
    
    # Import MainWindow
    from ui.main_window import MainWindow
    
    # Create a mock controller
    class MockController:
        def __init__(self):
            self.current_files = {}
    
    controller = MockController()
    
    # Create MainWindow instance
    main_window = MainWindow(controller)
    
    # Test that new variables are initialized
    print(f"master_valid_row_keys initialized: {hasattr(main_window, 'master_valid_row_keys')}")
    print(f"manual_invalid_row_keys initialized: {hasattr(main_window, 'manual_invalid_row_keys')}")
    print(f"is_comparison_mode initialized: {hasattr(main_window, 'is_comparison_mode')}")
    
    # Test initial values
    print(f"master_valid_row_keys type: {type(main_window.master_valid_row_keys)}")
    print(f"manual_invalid_row_keys type: {type(main_window.manual_invalid_row_keys)}")
    print(f"is_comparison_mode initial value: {main_window.is_comparison_mode}")
    
    # Test that they are properly initialized as sets
    assert isinstance(main_window.master_valid_row_keys, set), "master_valid_row_keys should be a set"
    assert isinstance(main_window.manual_invalid_row_keys, set), "manual_invalid_row_keys should be a set"
    assert isinstance(main_window.is_comparison_mode, bool), "is_comparison_mode should be a boolean"
    
    print("✅ All MainWindow integration tests passed!")

def test_row_classifier_import():
    """Test that the new row classifier functions can be imported and used"""
    print("\n=== Testing RowClassifier Import ===")
    
    from core.row_classifier import RowClassifier
    from utils.config import ColumnType
    
    classifier = RowClassifier()
    
    # Test that new methods exist
    print(f"validate_master_row_validity exists: {hasattr(classifier, 'validate_master_row_validity')}")
    print(f"validate_comparison_row_validity exists: {hasattr(classifier, 'validate_comparison_row_validity')}")
    print(f"_generate_row_key exists: {hasattr(classifier, '_generate_row_key')}")
    
    assert hasattr(classifier, 'validate_master_row_validity'), "validate_master_row_validity method missing"
    assert hasattr(classifier, 'validate_comparison_row_validity'), "validate_comparison_row_validity method missing"
    assert hasattr(classifier, '_generate_row_key'), "_generate_row_key method missing"
    
    print("✅ All RowClassifier import tests passed!")

if __name__ == "__main__":
    try:
        test_main_window_integration()
        test_row_classifier_import()
        print("\n🎉 Step 2 integration tests passed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc() 