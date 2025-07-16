#!/usr/bin/env python3
"""
Test script for new row validity functions
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.row_classifier import RowClassifier
from utils.config import ColumnType

def test_master_row_validity():
    """Test the master row validity function"""
    print("=== Testing Master Row Validity ===")
    
    classifier = RowClassifier()
    
    # Test case 1: Valid row with all required fields
    row_data = ["Item 1", "10", "100.50", "1005.00", "pcs", "001"]
    column_mapping = {
        0: ColumnType.DESCRIPTION,
        1: ColumnType.QUANTITY,
        2: ColumnType.UNIT_PRICE,
        3: ColumnType.TOTAL_PRICE,
        4: ColumnType.UNIT,
        5: ColumnType.CODE
    }
    
    result = classifier.validate_master_row_validity(row_data, column_mapping)
    print(f"Test 1 - Valid row: {result} (Expected: True)")
    assert result == True, "Valid row should return True"
    
    # Test case 2: Invalid row - missing description
    row_data = ["", "10", "100.50", "1005.00", "pcs", "001"]
    result = classifier.validate_master_row_validity(row_data, column_mapping)
    print(f"Test 2 - Missing description: {result} (Expected: False)")
    assert result == False, "Missing description should return False"
    
    # Test case 3: Invalid row - missing quantity
    row_data = ["Item 1", "", "100.50", "1005.00", "pcs", "001"]
    result = classifier.validate_master_row_validity(row_data, column_mapping)
    print(f"Test 3 - Missing quantity: {result} (Expected: False)")
    assert result == False, "Missing quantity should return False"
    
    # Test case 4: Valid row with zero values (should be allowed)
    row_data = ["Item 1", "0", "0", "0", "pcs", "001"]
    result = classifier.validate_master_row_validity(row_data, column_mapping)
    print(f"Test 4 - Zero values: {result} (Expected: True)")
    assert result == True, "Zero values should return True"
    
    print("✅ All master row validity tests passed!")

def test_comparison_row_validity():
    """Test the comparison row validity function"""
    print("\n=== Testing Comparison Row Validity ===")
    
    classifier = RowClassifier()
    
    # Sample row data
    row_data = ["Item 1", "10", "100.50", "1005.00", "pcs", "001"]
    column_mapping = {
        0: ColumnType.DESCRIPTION,
        1: ColumnType.QUANTITY,
        2: ColumnType.UNIT_PRICE,
        3: ColumnType.TOTAL_PRICE,
        4: ColumnType.UNIT,
        5: ColumnType.CODE
    }
    
    # Debug: Check what key is generated
    generated_key = classifier._generate_row_key(row_data, column_mapping)
    print(f"Generated key: '{generated_key}'")
    
    # Test case 1: Row is in master valid rows
    master_valid_rows = {generated_key, "Item 2||m2"}
    manual_invalid_rows = set()
    
    result = classifier.validate_comparison_row_validity(row_data, column_mapping, master_valid_rows, manual_invalid_rows)
    print(f"Test 1 - Row in master valid: {result} (Expected: True)")
    assert result == True, "Row in master valid should return True"
    
    # Test case 2: Row is manually marked as invalid
    master_valid_rows = {"Item 2||m2"}
    manual_invalid_rows = {generated_key}
    
    result = classifier.validate_comparison_row_validity(row_data, column_mapping, master_valid_rows, manual_invalid_rows)
    print(f"Test 2 - Manually invalid row: {result} (Expected: False)")
    assert result == False, "Manually invalid row should return False"
    
    # Test case 3: Row not in master but satisfies master criteria
    master_valid_rows = {"Item 2||m2"}
    manual_invalid_rows = set()
    
    result = classifier.validate_comparison_row_validity(row_data, column_mapping, master_valid_rows, manual_invalid_rows)
    print(f"Test 3 - New valid row: {result} (Expected: True)")
    assert result == True, "New valid row should return True"
    
    # Test case 4: Row not in master and doesn't satisfy master criteria
    invalid_row_data = ["Item 1", "", "100.50", "1005.00", "pcs", "001"]  # Missing quantity
    master_valid_rows = {"Item 2||m2"}
    manual_invalid_rows = set()
    
    result = classifier.validate_comparison_row_validity(invalid_row_data, column_mapping, master_valid_rows, manual_invalid_rows)
    print(f"Test 4 - New invalid row: {result} (Expected: False)")
    assert result == False, "New invalid row should return False"
    
    print("✅ All comparison row validity tests passed!")

def test_row_key_generation():
    """Test the row key generation function"""
    print("\n=== Testing Row Key Generation ===")
    
    classifier = RowClassifier()
    
    row_data = ["Item 1", "10", "100.50", "1005.00", "pcs", "001"]
    column_mapping = {
        0: ColumnType.DESCRIPTION,
        1: ColumnType.QUANTITY,
        2: ColumnType.UNIT_PRICE,
        3: ColumnType.TOTAL_PRICE,
        4: ColumnType.UNIT,
        5: ColumnType.CODE
    }
    
    key = classifier._generate_row_key(row_data, column_mapping)
    print(f"Generated key: '{key}'")
    expected_key = "Item 1|001|pcs"
    assert key == expected_key, f"Expected key '{expected_key}', got '{key}'"
    
    print("✅ Row key generation test passed!")

if __name__ == "__main__":
    try:
        test_master_row_validity()
        test_comparison_row_validity()
        test_row_key_generation()
        print("\n🎉 All tests passed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc() 