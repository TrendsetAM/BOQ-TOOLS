#!/usr/bin/env python3
"""
Comprehensive Test Script for Comparison BoQ Functions
Tests all the new functions implemented for the comparison logic replacement
"""

import sys
import os
import logging
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Any
import tempfile
import shutil

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent))

# Import the modules we need to test
from core.row_classifier import RowClassifier, InvalidRowsTracker, ROW_VALIDITY_STATIC, POSITION, calculate_cumulative_row_counts
from core.instance_matcher import InstanceMatcher, RowInstance, DatasetType
from core.comparison_engine import ComparisonEngine, MergeResult
from core.auto_categorizer import AutoCategorizer
from utils.config import ColumnType

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ComparisonFunctionsTester:
    """Comprehensive test suite for comparison functions"""
    
    def __init__(self):
        """Initialize the tester"""
        self.test_results = []
        self.passed_tests = 0
        self.failed_tests = 0
        
        # Create test data
        self._setup_test_data()
        
        # Initialize components
        self.row_classifier = RowClassifier()
        self.invalid_tracker = InvalidRowsTracker()
        self.instance_matcher = InstanceMatcher()
        self.comparison_engine = ComparisonEngine()
        
        # Create a temporary category dictionary for testing
        self._setup_test_category_dictionary()
        self.auto_categorizer = AutoCategorizer(self.test_category_dict)
    
    def _setup_test_data(self):
        """Set up test data for all tests"""
        # Test row data
        self.valid_row_data = [
            "Concrete Foundation", "10", "150.50", "1505.00", "m3", "CIV001"
        ]
        self.invalid_row_data = [
            "", "10", "150.50", "1505.00", "m3", "CIV001"  # Missing description
        ]
        self.invalid_row_data2 = [
            "Steel Beams", "", "150.50", "1505.00", "m3", "STEEL001"  # Missing quantity
        ]
        
        # Test column mapping
        self.column_mapping = {
            0: ColumnType.DESCRIPTION,
            1: ColumnType.QUANTITY,
            2: ColumnType.UNIT_PRICE,
            3: ColumnType.TOTAL_PRICE,
            4: ColumnType.UNIT,
            5: ColumnType.CODE
        }
        
        # Test sheets data
        self.sheets_data = {
            "Sheet1": [
                ["Concrete Foundation", "10", "150.50", "1505.00", "m3", "CIV001"],
                ["Steel Beams", "5", "200.00", "1000.00", "m", "STEEL001"],
                ["Electrical Wiring", "100", "25.00", "2500.00", "m", "ELEC001"]
            ],
            "Sheet2": [
                ["Solar Panels", "50", "300.00", "15000.00", "pcs", "SOLAR001"],
                ["Inverter", "2", "5000.00", "10000.00", "pcs", "INV001"]
            ]
        }
        
        # Test DataFrame
        self.test_dataframe = pd.DataFrame({
            'Description': ['Concrete Foundation', 'Steel Beams', 'Electrical Wiring'],
            'Quantity': [10, 5, 100],
            'Unit_Price': [150.50, 200.00, 25.00],
            'Total_Price': [1505.00, 1000.00, 2500.00],
            'Unit': ['m3', 'm', 'm'],
            'Code': ['CIV001', 'STEEL001', 'ELEC001'],
            'Category': ['Civil Works', '', 'Electrical Works']  # One empty category
        })
    
    def _setup_test_category_dictionary(self):
        """Set up a test category dictionary"""
        from core.category_dictionary import CategoryDictionary, CategoryMapping
        
        # Create a temporary directory for test dictionary
        self.test_dir = Path(tempfile.mkdtemp())
        self.test_dict_file = self.test_dir / "test_category_dictionary.json"
        
        # Create test mappings as CategoryMapping objects
        test_mappings = [
            CategoryMapping("concrete foundation", "Civil Works"),
            CategoryMapping("steel beams", "Civil Works"), 
            CategoryMapping("electrical wiring", "Electrical Works"),
            CategoryMapping("solar panels", "PV Mod. Installation"),
            CategoryMapping("inverter", "Electrical Works")
        ]
        
        # Save test dictionary
        import json
        from dataclasses import asdict
        
        data = {
            "mappings": [asdict(mapping) for mapping in test_mappings],
            "categories": list(set(mapping.category for mapping in test_mappings))
        }
        
        with open(self.test_dict_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        
        self.test_category_dict = CategoryDictionary(self.test_dict_file)
    
    def run_all_tests(self):
        """Run all tests and return results"""
        logger.info("Starting comprehensive test suite for comparison functions...")
        
        # Test ROW_VALIDITY function
        self._test_row_validity()
        
        # Test MANUAL_INVALID function
        self._test_manual_invalid()
        
        # Test MANUAL_OVERRIDE function
        self._test_manual_override()
        
        # Test POSITION function
        self._test_position()
        
        # Test LIST_INSTANCES function
        self._test_list_instances()
        
        # Test MERGE function
        self._test_merge()
        
        # Test ADD function
        self._test_add()
        
        # Test RECATEGORIZATION function
        self._test_recategorization()
        
        # Test integration scenarios
        self._test_integration_scenarios()
        
        # Print results
        self._print_results()
        
        # Cleanup
        self._cleanup()
        
        return self.passed_tests, self.failed_tests
    
    def _test_row_validity(self):
        """Test ROW_VALIDITY function"""
        logger.info("Testing ROW_VALIDITY function...")
        
        # Test valid row
        result = self.row_classifier.ROW_VALIDITY(self.valid_row_data, self.column_mapping)
        self._assert_test("ROW_VALIDITY - Valid row", result == True, f"Expected True, got {result}")
        
        # Test invalid row (missing description)
        result = self.row_classifier.ROW_VALIDITY(self.invalid_row_data, self.column_mapping)
        self._assert_test("ROW_VALIDITY - Invalid row (missing description)", result == False, f"Expected False, got {result}")
        
        # Test invalid row (missing quantity)
        result = self.row_classifier.ROW_VALIDITY(self.invalid_row_data2, self.column_mapping)
        self._assert_test("ROW_VALIDITY - Invalid row (missing quantity)", result == False, f"Expected False, got {result}")
        
        # Test static version
        result = ROW_VALIDITY_STATIC(self.valid_row_data, self.column_mapping)
        self._assert_test("ROW_VALIDITY_STATIC - Valid row", result == True, f"Expected True, got {result}")
    
    def _test_manual_invalid(self):
        """Test MANUAL_INVALID function"""
        logger.info("Testing MANUAL_INVALID function...")
        
        # Test adding manual invalidation
        result = self.invalid_tracker.MANUAL_INVALID("Test Description", 1, "Sheet1", "Test notes")
        self._assert_test("MANUAL_INVALID - Add invalidation", result == True, f"Expected True, got {result}")
        
        # Test adding duplicate (should fail)
        result = self.invalid_tracker.MANUAL_INVALID("Test Description", 1, "Sheet1", "Test notes")
        self._assert_test("MANUAL_INVALID - Duplicate invalidation", result == False, f"Expected False, got {result}")
        
        # Test adding different instance
        result = self.invalid_tracker.MANUAL_INVALID("Test Description", 2, "Sheet1", "Test notes")
        self._assert_test("MANUAL_INVALID - Different instance", result == True, f"Expected True, got {result}")
        
        # Test getting invalid rows set
        invalid_set = self.invalid_tracker.get_invalid_rows_set()
        expected_keys = {"Test Description|1", "Test Description|2"}
        self._assert_test("MANUAL_INVALID - Get invalid set", invalid_set == expected_keys, 
                         f"Expected {expected_keys}, got {invalid_set}")
        
        # Test count
        count = self.invalid_tracker.get_invalid_rows_count()
        self._assert_test("MANUAL_INVALID - Count", count == 2, f"Expected 2, got {count}")
    
    def _test_manual_override(self):
        """Test MANUAL_OVERRIDE function"""
        logger.info("Testing MANUAL_OVERRIDE function...")
        
        # Test checking existing invalidation
        result = self.invalid_tracker.MANUAL_OVERRIDE("Test Description", 1)
        self._assert_test("MANUAL_OVERRIDE - Existing invalidation", result == True, f"Expected True, got {result}")
        
        # Test checking non-existing invalidation
        result = self.invalid_tracker.MANUAL_OVERRIDE("Non-existent Description", 1)
        self._assert_test("MANUAL_OVERRIDE - Non-existing invalidation", result == False, f"Expected False, got {result}")
        
        # Test checking different instance
        result = self.invalid_tracker.MANUAL_OVERRIDE("Test Description", 3)
        self._assert_test("MANUAL_OVERRIDE - Different instance", result == False, f"Expected False, got {result}")
    
    def _test_position(self):
        """Test POSITION function"""
        logger.info("Testing POSITION function...")
        
        # Test basic position calculation
        position = POSITION("Sheet1", 5, 1, 0)
        self._assert_test("POSITION - Basic calculation", position == 6, f"Expected 6, got {position}")
        
        # Test with cumulative offset
        position = POSITION("Sheet2", 3, 2, 10)
        self._assert_test("POSITION - With offset", position == 14, f"Expected 14, got {position}")
        
        # Test cumulative row counts calculation
        cumulative_counts = calculate_cumulative_row_counts(self.sheets_data, {
            "Sheet1": self.column_mapping,
            "Sheet2": self.column_mapping
        })
        expected_counts = {"Sheet1": 0, "Sheet2": 3}  # 3 valid rows in Sheet1
        self._assert_test("POSITION - Cumulative counts", cumulative_counts == expected_counts, 
                         f"Expected {expected_counts}, got {cumulative_counts}")
    
    def _test_list_instances(self):
        """Test LIST_INSTANCES function"""
        logger.info("Testing LIST_INSTANCES function...")
        
        # Create test row instances
        comparison_instances = [
            RowInstance(1, ["Concrete Foundation", "10", "150.50"], "Sheet1", 0, "Concrete Foundation", 1),
            RowInstance(2, ["Steel Beams", "5", "200.00"], "Sheet1", 1, "Steel Beams", 1),
            RowInstance(3, ["Concrete Foundation", "15", "160.00"], "Sheet2", 0, "Concrete Foundation", 2)
        ]
        
        dataset_instances = [
            RowInstance(1, ["Concrete Foundation", "8", "140.00"], "Dataset", 0, "Concrete Foundation", 1),
            RowInstance(2, ["Steel Beams", "3", "180.00"], "Dataset", 1, "Steel Beams", 1)
        ]
        
        # Test finding instances in comparison BoQ
        result = self.instance_matcher.get_comparison_instances(1, "Concrete Foundation", comparison_instances)
        self._assert_test("LIST_INSTANCES - Comparison instances", len(result) == 2, f"Expected 2, got {len(result)}")
        
        # Test finding instances in dataset
        result = self.instance_matcher.get_dataset_instances(1, "Concrete Foundation", dataset_instances)
        self._assert_test("LIST_INSTANCES - Dataset instances", len(result) == 1, f"Expected 1, got {len(result)}")
        
        # Test validation
        result = self.instance_matcher.validate_instance_count(
            comparison_instances[:2], dataset_instances[:1], "Concrete Foundation"
        )
        self._assert_test("LIST_INSTANCES - Validation", result == True, f"Expected True, got {result}")
    
    def _test_merge(self):
        """Test MERGE function"""
        logger.info("Testing MERGE function...")
        
        # Create test DataFrame
        df = self.test_dataframe.copy()
        
        # Test merging data
        comparison_row = ["Concrete Foundation", "12", "160.00", "1920.00", "m3", "CIV001"]
        result = self.comparison_engine.MERGE(comparison_row, df, "Offer1", self.column_mapping, 0)
        
        self._assert_test("MERGE - Success", result.success == True, f"Expected True, got {result.success}")
        self._assert_test("MERGE - Rows updated", result.rows_updated == 1, f"Expected 1, got {result.rows_updated}")
        self._assert_test("MERGE - No errors", len(result.errors) == 0, f"Expected 0 errors, got {len(result.errors)}")
        
        # Check if offer columns were created
        expected_columns = ['quantity[Offer1]', 'unit_price[Offer1]', 'total_price[Offer1]', 
                          'manhours[Offer1]', 'wage[Offer1]']
        for col in expected_columns:
            self._assert_test(f"MERGE - Column {col} created", col in df.columns, f"Column {col} not found")
        
        # Test validation
        result = self.comparison_engine.validate_merge_operation(df, "Offer1", 0)
        self._assert_test("MERGE - Validation", result == True, f"Expected True, got {result}")
    
    def _test_add(self):
        """Test ADD function"""
        logger.info("Testing ADD function...")
        
        # Create test DataFrame
        df = self.test_dataframe.copy()
        initial_length = len(df)
        
        # Test adding new row
        comparison_row = ["New Item", "20", "100.00", "2000.00", "pcs", "NEW001"]
        result = self.comparison_engine.ADD(comparison_row, df, self.column_mapping, 100)
        
        self._assert_test("ADD - Success", result["success"] == True, f"Expected True, got {result['success']}")
        self._assert_test("ADD - Row added", result["row_added"] == True, f"Expected True, got {result['row_added']}")
        self._assert_test("ADD - Length increased", len(df) == initial_length + 1, 
                         f"Expected {initial_length + 1}, got {len(df)}")
        
        # Check if new row has correct data
        new_row = df.iloc[-1]
        self._assert_test("ADD - Description", new_row['Description'] == "New Item", 
                         f"Expected 'New Item', got '{new_row['Description']}'")
        self._assert_test("ADD - Position", new_row['Position'] == 100, 
                         f"Expected 100, got {new_row['Position']}")
    
    def _test_recategorization(self):
        """Test RECATEGORIZATION function"""
        logger.info("Testing RECATEGORIZATION function...")
        
        # Create DataFrame with empty categories
        df = self.test_dataframe.copy()
        df.loc[1, 'Category'] = ''  # Make sure Steel Beams has empty category
        
        # Test recategorization
        result = self.auto_categorizer.RECATEGORIZATION(df, 'Description', 'Category')
        
        self._assert_test("RECATEGORIZATION - Success", result.match_rate > 0, 
                         f"Expected match_rate > 0, got {result.match_rate}")
        self._assert_test("RECATEGORIZATION - DataFrame returned", result.dataframe is not None, 
                         "DataFrame is None")
        
        # Check if empty categories were filled
        empty_categories = result.dataframe[result.dataframe['Category'].isna() | (result.dataframe['Category'] == '')]
        self._assert_test("RECATEGORIZATION - Empty categories reduced", len(empty_categories) < 2, 
                         f"Expected fewer empty categories, got {len(empty_categories)}")
    
    def _test_integration_scenarios(self):
        """Test integration scenarios"""
        logger.info("Testing integration scenarios...")
        
        # Scenario 1: Complete workflow with valid data
        self._test_complete_workflow()
        
        # Scenario 2: Error handling
        self._test_error_handling()
        
        # Scenario 3: Edge cases
        self._test_edge_cases()
    
    def _test_complete_workflow(self):
        """Test a complete workflow scenario"""
        logger.info("Testing complete workflow...")
        
        # 1. Validate rows
        valid_rows = []
        for row_data in self.sheets_data["Sheet1"]:
            if self.row_classifier.ROW_VALIDITY(row_data, self.column_mapping):
                valid_rows.append(row_data)
        
        self._assert_test("Workflow - Valid rows found", len(valid_rows) > 0, 
                         f"Expected > 0 valid rows, got {len(valid_rows)}")
        
        # 2. Add manual invalidations
        self.invalid_tracker.MANUAL_INVALID("Concrete Foundation", 1, "Sheet1", "Test")
        
        # 3. Check manual overrides
        should_invalidate = self.invalid_tracker.MANUAL_OVERRIDE("Concrete Foundation", 1)
        self._assert_test("Workflow - Manual override", should_invalidate == True, 
                         f"Expected True, got {should_invalidate}")
        
        # 4. Calculate positions
        position = POSITION("Sheet1", 1, 1, 0)
        self._assert_test("Workflow - Position calculation", position > 0, 
                         f"Expected > 0, got {position}")
        
        # 5. Merge data
        df = self.test_dataframe.copy()
        result = self.comparison_engine.MERGE(self.valid_row_data, df, "TestOffer", self.column_mapping, 0)
        self._assert_test("Workflow - Merge success", result.success == True, 
                         f"Expected True, got {result.success}")
    
    def _test_error_handling(self):
        """Test error handling"""
        logger.info("Testing error handling...")
        
        # Test with empty data
        result = self.row_classifier.ROW_VALIDITY([], self.column_mapping)
        self._assert_test("Error handling - Empty data", result == False, f"Expected False, got {result}")
        
        # Test with invalid column mapping
        result = self.row_classifier.ROW_VALIDITY(self.valid_row_data, {})
        self._assert_test("Error handling - Empty mapping", result == False, f"Expected False, got {result}")
    
    def _test_edge_cases(self):
        """Test edge cases"""
        logger.info("Testing edge cases...")
        
        # Test with European number formats
        european_row = ["Test Item", "1.234,56", "2.345,67", "2.890,12", "m3", "TEST001"]
        result = self.row_classifier.ROW_VALIDITY(european_row, self.column_mapping)
        self._assert_test("Edge cases - European format", result == True, f"Expected True, got {result}")
        
        # Test with currency symbols
        currency_row = ["Test Item", "1,234.56", "$2,345.67", "€2,890.12", "m3", "TEST001"]
        result = self.row_classifier.ROW_VALIDITY(currency_row, self.column_mapping)
        self._assert_test("Edge cases - Currency symbols", result == True, f"Expected True, got {result}")
        
        # Test with zero values
        zero_row = ["Test Item", "0", "0.00", "0.00", "m3", "TEST001"]
        result = self.row_classifier.ROW_VALIDITY(zero_row, self.column_mapping)
        self._assert_test("Edge cases - Zero values", result == True, f"Expected True, got {result}")
    
    def _assert_test(self, test_name: str, condition: bool, message: str):
        """Assert a test condition and record the result"""
        if condition:
            self.passed_tests += 1
            logger.info(f"✅ PASS: {test_name}")
        else:
            self.failed_tests += 1
            logger.error(f"❌ FAIL: {test_name} - {message}")
        
        self.test_results.append({
            'name': test_name,
            'passed': condition,
            'message': message if not condition else "PASS"
        })
    
    def _print_results(self):
        """Print test results summary"""
        logger.info("\n" + "="*60)
        logger.info("TEST RESULTS SUMMARY")
        logger.info("="*60)
        logger.info(f"Total Tests: {self.passed_tests + self.failed_tests}")
        logger.info(f"Passed: {self.passed_tests}")
        logger.info(f"Failed: {self.failed_tests}")
        logger.info(f"Success Rate: {(self.passed_tests / (self.passed_tests + self.failed_tests) * 100):.1f}%")
        
        if self.failed_tests > 0:
            logger.info("\nFailed Tests:")
            for result in self.test_results:
                if not result['passed']:
                    logger.error(f"  - {result['name']}: {result['message']}")
        
        logger.info("="*60)
    
    def _cleanup(self):
        """Clean up test resources"""
        try:
            if hasattr(self, 'test_dir') and self.test_dir.exists():
                shutil.rmtree(self.test_dir)
                logger.info("Test cleanup completed")
        except Exception as e:
            logger.warning(f"Cleanup warning: {e}")


def main():
    """Main test runner"""
    logger.info("Starting comprehensive test suite for comparison functions...")
    
    tester = ComparisonFunctionsTester()
    passed, failed = tester.run_all_tests()
    
    if failed == 0:
        logger.info("🎉 All tests passed! The comparison functions are working correctly.")
        return 0
    else:
        logger.error(f"❌ {failed} tests failed. Please review the implementation.")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code) 