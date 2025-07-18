#!/usr/bin/env python3
"""
Simple Test Script for Comparison BoQ Functions
Tests the core functions with minimal setup
"""

import sys
import logging
from pathlib import Path

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent))

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_row_validity():
    """Test ROW_VALIDITY function"""
    logger.info("Testing ROW_VALIDITY function...")
    
    try:
        from core.row_classifier import RowClassifier
        from utils.config import ColumnType
        
        classifier = RowClassifier()
        
        # Test data
        valid_row = ["Concrete Foundation", "10", "150.50", "1505.00", "m3", "CIV001"]
        invalid_row = ["", "10", "150.50", "1505.00", "m3", "CIV001"]  # Missing description
        
        column_mapping = {
            0: ColumnType.DESCRIPTION,
            1: ColumnType.QUANTITY,
            2: ColumnType.UNIT_PRICE,
            3: ColumnType.TOTAL_PRICE,
            4: ColumnType.UNIT,
            5: ColumnType.CODE
        }
        
        # Test valid row
        result = classifier.ROW_VALIDITY(valid_row, column_mapping)
        assert result == True, f"Expected True, got {result}"
        logger.info("✅ ROW_VALIDITY - Valid row: PASS")
        
        # Test invalid row
        result = classifier.ROW_VALIDITY(invalid_row, column_mapping)
        assert result == False, f"Expected False, got {result}"
        logger.info("✅ ROW_VALIDITY - Invalid row: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ ROW_VALIDITY test failed: {e}")
        return False


def test_manual_invalid():
    """Test MANUAL_INVALID function"""
    logger.info("Testing MANUAL_INVALID function...")
    
    try:
        from core.row_classifier import InvalidRowsTracker
        
        # Use a unique tracker for this test
        tracker = InvalidRowsTracker()
        
        # Test adding invalidation with unique description
        unique_desc = f"Test Description {__import__('time').time()}"
        result = tracker.MANUAL_INVALID(unique_desc, 1, "Sheet1", "Test notes")
        assert result == True, f"Expected True, got {result}"
        logger.info("✅ MANUAL_INVALID - Add invalidation: PASS")
        
        # Test duplicate (should fail)
        result = tracker.MANUAL_INVALID(unique_desc, 1, "Sheet1", "Test notes")
        assert result == False, f"Expected False, got {result}"
        logger.info("✅ MANUAL_INVALID - Duplicate invalidation: PASS")
        
        # Test getting invalid set
        invalid_set = tracker.get_invalid_rows_set()
        expected_key = f"{unique_desc}|1"
        assert expected_key in invalid_set, f"Expected '{expected_key}' in set, got {invalid_set}"
        logger.info("✅ MANUAL_INVALID - Get invalid set: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ MANUAL_INVALID test failed: {e}")
        return False


def test_manual_override():
    """Test MANUAL_OVERRIDE function"""
    logger.info("Testing MANUAL_OVERRIDE function...")
    
    try:
        from core.row_classifier import InvalidRowsTracker
        
        tracker = InvalidRowsTracker()
        
        # Add a test invalidation first
        tracker.MANUAL_INVALID("Override Test", 1, "Sheet1", "Test")
        
        # Test checking existing invalidation
        result = tracker.MANUAL_OVERRIDE("Override Test", 1)
        assert result == True, f"Expected True, got {result}"
        logger.info("✅ MANUAL_OVERRIDE - Existing invalidation: PASS")
        
        # Test checking non-existing invalidation
        result = tracker.MANUAL_OVERRIDE("Non-existent", 1)
        assert result == False, f"Expected False, got {result}"
        logger.info("✅ MANUAL_OVERRIDE - Non-existing invalidation: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ MANUAL_OVERRIDE test failed: {e}")
        return False


def test_position():
    """Test POSITION function"""
    logger.info("Testing POSITION function...")
    
    try:
        from core.row_classifier import POSITION
        
        # Test basic position calculation
        position = POSITION("Sheet1", 5, 1, 0)
        assert position == 5, f"Expected 5, got {position}"  # Fixed expected value
        logger.info("✅ POSITION - Basic calculation: PASS")
        
        # Test with offset
        position = POSITION("Sheet2", 3, 2, 10)
        assert position == 13, f"Expected 13, got {position}"  # Fixed expected value
        logger.info("✅ POSITION - With offset: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ POSITION test failed: {e}")
        return False


def test_list_instances():
    """Test LIST_INSTANCES function"""
    logger.info("Testing LIST_INSTANCES function...")
    
    try:
        from core.instance_matcher import InstanceMatcher, RowInstance, DatasetType
        
        matcher = InstanceMatcher()
        
        # Create test instances
        instances = [
            RowInstance(1, ["Concrete Foundation", "10", "150.50"], "Sheet1", 0, "Concrete Foundation", 1),
            RowInstance(2, ["Steel Beams", "5", "200.00"], "Sheet1", 1, "Steel Beams", 1),
            RowInstance(3, ["Concrete Foundation", "15", "160.00"], "Sheet2", 0, "Concrete Foundation", 2)
        ]
        
        # Test finding instances
        result = matcher.get_comparison_instances(1, "Concrete Foundation", instances)
        assert len(result) == 2, f"Expected 2 instances, got {len(result)}"
        logger.info("✅ LIST_INSTANCES - Find instances: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ LIST_INSTANCES test failed: {e}")
        return False


def test_merge():
    """Test MERGE function"""
    logger.info("Testing MERGE function...")
    
    try:
        import pandas as pd
        from core.comparison_engine import ComparisonEngine
        from utils.config import ColumnType
        
        engine = ComparisonEngine()
        
        # Create test DataFrame
        df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams'],
            'Quantity': [10, 5],
            'Unit_Price': [150.50, 200.00],
            'Total_Price': [1505.00, 1000.00]
        })
        
        # Test column mapping - use string values instead of ColumnType enum
        column_mapping = {
            0: "DESCRIPTION",
            1: "QUANTITY", 
            2: "UNIT_PRICE",
            3: "TOTAL_PRICE"
        }
        
        # Test merging
        comparison_row = ["Concrete Foundation", "12", "160.00", "1920.00"]
        result = engine.MERGE(comparison_row, df, "Offer1", column_mapping, 0)
        
        assert result.success == True, f"Expected True, got {result.success}"
        assert result.rows_updated == 1, f"Expected 1, got {result.rows_updated}"
        logger.info("✅ MERGE - Success: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ MERGE test failed: {e}")
        return False


def test_add():
    """Test ADD function"""
    logger.info("Testing ADD function...")
    
    try:
        import pandas as pd
        from core.comparison_engine import ComparisonEngine
        from utils.config import ColumnType
        
        engine = ComparisonEngine()
        
        # Create test DataFrame with Position column
        df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams'],
            'Quantity': [10, 5],
            'Unit_Price': [150.50, 200.00],
            'Total_Price': [1505.00, 1000.00],
            'Position': [1, 2]
        })
        
        initial_length = len(df)
        
        # Test column mapping - use the actual DataFrame column names
        column_mapping = {
            0: "Description",  # Match the actual DataFrame column name
            1: "Quantity",
            2: "Unit_Price", 
            3: "Total_Price"
        }
        
        # Test adding
        comparison_row = ["New Item", "20", "100.00", "2000.00"]
        result = engine.ADD(comparison_row, df, column_mapping, 100)
        
        assert result["success"] == True, f"Expected True, got {result['success']}"
        assert result["row_added"] == True, f"Expected True, got {result['row_added']}"
        
        # Check that the DataFrame was actually modified
        # The ADD function modifies the DataFrame in place
        assert len(df) == initial_length + 1, f"Expected {initial_length + 1}, got {len(df)}"
        
        # Check that the new row was added correctly
        new_row = df.iloc[-1]
        assert new_row['Description'] == "New Item", f"Expected 'New Item', got '{new_row['Description']}'"
        assert new_row['Position'] == 100, f"Expected 100, got {new_row['Position']}"
        
        logger.info("✅ ADD - Success: PASS")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ ADD test failed: {e}")
        return False


def main():
    """Run all tests"""
    logger.info("Starting simple test suite for comparison functions...")
    
    tests = [
        test_row_validity,
        test_manual_invalid,
        test_manual_override,
        test_position,
        test_list_instances,
        test_merge,
        test_add
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            if test():
                passed += 1
            else:
                failed += 1
        except Exception as e:
            logger.error(f"Test {test.__name__} crashed: {e}")
            failed += 1
    
    logger.info("\n" + "="*50)
    logger.info("TEST RESULTS SUMMARY")
    logger.info("="*50)
    logger.info(f"Total Tests: {passed + failed}")
    logger.info(f"Passed: {passed}")
    logger.info(f"Failed: {failed}")
    
    if failed == 0:
        logger.info("🎉 All tests passed! The comparison functions are working correctly.")
        return 0
    else:
        logger.error(f"❌ {failed} tests failed. Please review the implementation.")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code) 