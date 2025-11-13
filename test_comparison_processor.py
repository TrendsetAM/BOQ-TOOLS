#!/usr/bin/env python3
"""
Comprehensive Test Script for Comparison BoQ Functions and ComparisonProcessor
Tests all functionalities from both phases including core functions and the new ComparisonProcessor class
"""

import sys
import logging
import pandas as pd
from pathlib import Path

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent))

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_comparison_processor_initialization():
    """Test ComparisonProcessor initialization"""
    logger.info("Testing ComparisonProcessor initialization...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Check initial state
        assert processor.master_dataset is None, "master_dataset should be None initially"
        assert processor.comparison_data is None, "comparison_data should be None initially"
        assert processor.manual_invalidations == set(), "manual_invalidations should be empty set"
        assert processor.row_results == [], "row_results should be empty list"
        assert processor.instance_matches == [], "instance_matches should be empty list"
        assert processor.merge_results == [], "merge_results should be empty list"
        assert processor.add_results == [], "add_results should be empty list"
        
        logger.info("✅ ComparisonProcessor initialization: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ ComparisonProcessor initialization test failed: {e}")
        return False


def test_data_loading():
    """Test data loading methods"""
    logger.info("Testing data loading methods...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Create test DataFrames
        master_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams', 'Electrical Wiring'],
            'Quantity': [10, 5, 100],
            'Unit_Price': [150.50, 200.00, 25.00],
            'Total_Price': [1505.00, 1000.00, 2500.00],
            'Category': ['Civil', 'Structural', 'Electrical']
        })
        
        comparison_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'New Item', 'Steel Beams'],
            'Quantity': [12, 20, 8],
            'Unit_Price': [160.00, 100.00, 220.00],
            'Total_Price': [1920.00, 2000.00, 1760.00],
            'Category': ['Civil', '', 'Structural']
        })
        
        # Test loading master dataset
        manual_invalidations = {'Concrete Foundation|1', 'Steel Beams|1'}
        processor.load_master_dataset(master_df, manual_invalidations)
        
        assert processor.master_dataset is not None, "master_dataset should be loaded"
        assert len(processor.master_dataset) == 3, f"Expected 3 rows, got {len(processor.master_dataset)}"
        assert processor.manual_invalidations == {'Concrete Foundation|1', 'Steel Beams|1'}, "manual_invalidations should be set"
        
        # Test loading comparison data
        processor.load_comparison_data(comparison_df)
        
        assert processor.comparison_data is not None, "comparison_data should be loaded"
        assert len(processor.comparison_data) == 3, f"Expected 3 rows, got {len(processor.comparison_data)}"
        
        logger.info("✅ Data loading methods: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Data loading test failed: {e}")
        return False


def test_validation():
    """Test data validation"""
    logger.info("Testing data validation...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Test validation with no data loaded
        is_valid, message = processor.validate_comparison_data()
        assert is_valid == False, f"Expected False, got {is_valid}"
        assert "not loaded" in message, f"Expected error message about not loaded, got {message}"
        
        # Test validation with compatible data
        master_df = pd.DataFrame({
            'Description': ['Item 1'],
            'Quantity': [10],
            'Unit_Price': [100.00]
        })
        
        comparison_df = pd.DataFrame({
            'Description': ['Item 2'],
            'Quantity': [5],
            'Unit_Price': [50.00]
        })
        
        processor.load_master_dataset(master_df)
        processor.load_comparison_data(comparison_df)
        
        is_valid, message = processor.validate_comparison_data()
        assert is_valid == True, f"Expected True, got {is_valid}"
        assert "successful" in message, f"Expected success message, got {message}"
        
        # Test validation with incompatible data
        incompatible_df = pd.DataFrame({
            'Description': ['Item 3'],
            'Quantity': [15]
            # Missing Unit_Price column
        })
        
        processor.load_comparison_data(incompatible_df)
        
        is_valid, message = processor.validate_comparison_data()
        assert is_valid == False, f"Expected False, got {is_valid}"
        assert "missing columns" in message, f"Expected error about missing columns, got {message}"
        
        logger.info("✅ Data validation: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Data validation test failed: {e}")
        return False


def test_row_processing():
    """Test row processing logic"""
    logger.info("Testing row processing logic...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Create test data with manual invalidations
        master_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams'],
            'Quantity': [10, 5],
            'Unit_Price': [150.50, 200.00],
            'Total_Price': [1505.00, 1000.00]
        })
        
        comparison_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams', 'New Item'],
            'Quantity': [12, 8, 20],
            'Unit_Price': [160.00, 220.00, 100.00],
            'Total_Price': [1920.00, 1760.00, 2000.00]
        })
        
        # Load data with manual invalidations
        manual_invalidations = {'Concrete Foundation|1'}  # Only invalidate one item
        processor.load_master_dataset(master_df, manual_invalidations)
        processor.load_comparison_data(comparison_df)
        
        # Process rows
        results = processor.process_comparison_rows()
        
        assert len(results) == 3, f"Expected 3 results, got {len(results)}"
        
        # Check manual override results
        manual_override_results = [r for r in results if r['reason'] == 'MANUAL_OVERRIDE']
        assert len(manual_override_results) == 1, f"Expected 1 manual override, got {len(manual_override_results)}"
        
        # Check row validity results
        validity_results = [r for r in results if r['reason'] == 'ROW_VALIDITY']
        assert len(validity_results) == 2, f"Expected 2 validity checks, got {len(validity_results)}"
        
        # Check that the valid rows are Steel Beams and New Item
        valid_rows = [r for r in results if r['is_valid'] and r['reason'] == 'ROW_VALIDITY']
        assert len(valid_rows) == 2, f"Expected 2 valid rows, got {len(valid_rows)}"
        
        logger.info("✅ Row processing logic: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Row processing test failed: {e}")
        return False


def test_instance_matching():
    """Test instance matching and merge/add logic"""
    logger.info("Testing instance matching and merge/add logic...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Create test data
        master_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams'],
            'Quantity': [10, 5],
            'Unit_Price': [150.50, 200.00],
            'Total_Price': [1505.00, 1000.00],
            'Position': [1, 2]
        })
        
        comparison_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams', 'New Item'],
            'Quantity': [12, 8, 20],
            'Unit_Price': [160.00, 220.00, 100.00],
            'Total_Price': [1920.00, 1760.00, 2000.00],
            'Position': [1, 2, 3]
        })
        
        # Load data
        processor.load_master_dataset(master_df)
        processor.load_comparison_data(comparison_df)
        
        # Process rows first
        processor.process_comparison_rows()
        
        # Process valid rows
        results = processor.process_valid_rows()
        
        # Debug: check what happened
        logger.info(f"Results: {results}")
        logger.info(f"Master dataset columns: {list(processor.master_dataset.columns)}")
        logger.info(f"Master dataset shape: {processor.master_dataset.shape}")
        
        # Should have 2 merges (Concrete Foundation and Steel Beams) and 1 add (New Item)
        merge_ops = [r for r in results if r['type'] == 'MERGE']
        add_ops = [r for r in results if r['type'] == 'ADD']
        
        logger.info(f"Merge operations: {len(merge_ops)}")
        logger.info(f"Add operations: {len(add_ops)}")
        
        assert len(merge_ops) == 2, f"Expected 2 merge operations, got {len(merge_ops)}"
        assert len(add_ops) == 1, f"Expected 1 add operation, got {len(add_ops)}"
        
        # Check that the master dataset was modified
        assert len(processor.master_dataset) == 3, f"Expected 3 rows in master dataset, got {len(processor.master_dataset)}"
        
        # Debug: check what columns exist
        logger.info(f"Master dataset columns: {list(processor.master_dataset.columns)}")
        logger.info(f"Master dataset shape: {processor.master_dataset.shape}")
        
        logger.info("✅ Instance matching and merge/add logic: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Instance matching test failed: {e}")
        return False


def test_data_cleanup():
    """Test data cleanup functionality"""
    logger.info("Testing data cleanup functionality...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Create test data with empty values
        comparison_df = pd.DataFrame({
            'Description': ['Item 1', 'Item 2', 'Item 3'],
            'Quantity': [10, 5, 15],
            'Unit_Price': [100.00, '', 75.00],  # Empty value
            'Total_Price': [1000.00, 500.00, ''],  # Empty value
            'Category': ['Civil', '', 'Electrical']  # Empty category
        })
        
        processor.load_comparison_data(comparison_df)
        
        # Test cleanup
        recat_results = processor.cleanup_comparison_data()
        
        # Check that empty numeric values were replaced with 0
        assert processor.comparison_data['Unit_Price'].iloc[1] == 0, "Empty Unit_Price should be 0"
        assert processor.comparison_data['Total_Price'].iloc[2] == 0, "Empty Total_Price should be 0"
        
        # Check that non-empty values remain unchanged
        assert processor.comparison_data['Unit_Price'].iloc[0] == 100.00, "Non-empty Unit_Price should remain unchanged"
        assert processor.comparison_data['Total_Price'].iloc[0] == 1000.00, "Non-empty Total_Price should remain unchanged"
        
        logger.info("✅ Data cleanup functionality: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Data cleanup test failed: {e}")
        return False


def test_integration_workflow():
    """Test the complete integration workflow"""
    logger.info("Testing complete integration workflow...")
    
    try:
        from core.comparison_engine import ComparisonProcessor
        
        processor = ComparisonProcessor()
        
        # Create comprehensive test data
        master_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams', 'Electrical Wiring'],
            'Quantity': [10, 5, 100],
            'Unit_Price': [150.50, 200.00, 25.00],
            'Total_Price': [1505.00, 1000.00, 2500.00],
            'Category': ['Civil', 'Structural', 'Electrical'],
            'Position': [1, 2, 3]
        })
        
        comparison_df = pd.DataFrame({
            'Description': ['Concrete Foundation', 'New Item', 'Steel Beams', 'Another Item'],
            'Quantity': [12, 20, 8, 30],
            'Unit_Price': [160.00, 100.00, 220.00, ''],
            'Total_Price': [1920.00, 2000.00, 1760.00, ''],
            'Category': ['Civil', '', 'Structural', ''],
            'Position': [1, 2, 3, 4]
        })
        
        # Load data with manual invalidations
        manual_invalidations = {'Concrete Foundation|1'}
        processor.load_master_dataset(master_df, manual_invalidations)
        processor.load_comparison_data(comparison_df)
        
        # Validate data
        is_valid, message = processor.validate_comparison_data()
        assert is_valid == True, f"Data validation failed: {message}"
        
        # Process rows
        row_results = processor.process_comparison_rows()
        assert len(row_results) == 4, f"Expected 4 row results, got {len(row_results)}"
        
        # Check manual override
        manual_overrides = [r for r in row_results if r['reason'] == 'MANUAL_OVERRIDE']
        assert len(manual_overrides) == 1, f"Expected 1 manual override, got {len(manual_overrides)}"
        
        # Process valid rows
        operation_results = processor.process_valid_rows()
        assert len(operation_results) > 0, "Expected operation results"
        
        # Cleanup data
        cleanup_results = processor.cleanup_comparison_data()
        
        # Verify final state
        assert processor.master_dataset is not None, "Master dataset should exist"
        assert processor.comparison_data is not None, "Comparison data should exist"
        assert len(processor.row_results) == 4, "Should have 4 row results"
        
        logger.info("✅ Complete integration workflow: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Integration workflow test failed: {e}")
        return False


def test_core_functions_integration():
    """Test that all core functions work with ComparisonProcessor"""
    logger.info("Testing core functions integration...")
    
    try:
        from core.comparison_engine import ComparisonProcessor, ComparisonEngine
        from core.row_classifier import RowClassifier, InvalidRowsTracker
        from core.instance_matcher import InstanceMatcher
        
        # Test ROW_VALIDITY integration
        classifier = RowClassifier()
        valid_row = ["Concrete Foundation", "10", "150.50", "1505.00"]
        invalid_row = ["", "10", "150.50", "1505.00"]
        
        from utils.config import ColumnType
        column_mapping = {0: ColumnType.DESCRIPTION, 1: ColumnType.QUANTITY, 2: ColumnType.UNIT_PRICE, 3: ColumnType.TOTAL_PRICE}
        
        # Convert to string values as expected by ROW_VALIDITY
        valid_row_str = [str(val) for val in valid_row]
        invalid_row_str = [str(val) for val in invalid_row]
        
        assert classifier.ROW_VALIDITY(valid_row_str, column_mapping) == True, "Valid row should pass"
        assert classifier.ROW_VALIDITY(invalid_row_str, column_mapping) == False, "Invalid row should fail"
        
        # Test MANUAL_INVALID integration
        tracker = InvalidRowsTracker()
        unique_desc = f"Test Description {__import__('time').time()}"
        result = tracker.MANUAL_INVALID(unique_desc, 1, "Sheet1", "Test notes")
        assert result == True, "Manual invalidation should succeed"
        
        # Test MANUAL_OVERRIDE integration
        result = tracker.MANUAL_OVERRIDE(unique_desc, 1)
        assert result == True, "Manual override should find existing invalidation"
        
        # Test POSITION integration
        from core.row_classifier import POSITION
        position = POSITION("Sheet1", 5, 1, 0)
        assert position == 5, f"Expected position 5, got {position}"
        
        # Test MERGE integration
        engine = ComparisonEngine()
        df = pd.DataFrame({
            'Description': ['Concrete Foundation'],
            'Quantity': [10],
            'Unit_Price': [150.50],
            'Total_Price': [1505.00]
        })
        
        comparison_row = ["Concrete Foundation", "12", "160.00", "1920.00"]
        merge_result = engine.MERGE(comparison_row, df, "Offer1", column_mapping, 0)
        assert merge_result.success == True, "Merge should succeed"
        
        # Test ADD integration
        add_result = engine.ADD(comparison_row, df, column_mapping, 100)
        assert add_result["success"] == True, "Add should succeed"
        assert add_result["row_added"] == True, "Row should be added"
        
        logger.info("✅ Core functions integration: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Core functions integration test failed: {e}")
        return False


def main():
    """Run all tests"""
    logger.info("Starting comprehensive test suite for ComparisonProcessor and core functions...")
    
    tests = [
        test_comparison_processor_initialization,
        test_data_loading,
        test_validation,
        test_row_processing,
        test_instance_matching,
        test_data_cleanup,
        test_integration_workflow,
        test_core_functions_integration
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
    
    logger.info("\n" + "="*60)
    logger.info("COMPREHENSIVE TEST RESULTS SUMMARY")
    logger.info("="*60)
    logger.info(f"Total Tests: {passed + failed}")
    logger.info(f"Passed: {passed}")
    logger.info(f"Failed: {failed}")
    
    if failed == 0:
        logger.info("🎉 All tests passed! The ComparisonProcessor and core functions are working correctly.")
        return 0
    else:
        logger.error(f"❌ {failed} tests failed. Please review the implementation.")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code) 