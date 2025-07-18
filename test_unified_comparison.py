#!/usr/bin/env python3
"""
Test script for unified comparison approach
"""

import sys
import os
import logging
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from ui.main_window import MainWindow
from core.comparison_engine import ComparisonProcessor
import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(name)s: %(message)s')
logger = logging.getLogger(__name__)

def test_unified_dataframe_creation():
    """Test the unified DataFrame creation method"""
    logger.info("Testing unified DataFrame creation...")
    
    # Create a mock file mapping with dataframe attribute
    class MockFileMapping:
        def __init__(self):
            self.dataframe = pd.DataFrame({
                'Description': ['Test Description 1', 'Test Description 2', 'Test Description 3'],
                'code': ['CODE1', 'CODE2', 'CODE3'],
                'unit': ['EA', 'EA', 'EA'],
                'quantity': [1, 2, 3],
                'unit_price': [100, 200, 300],
                'total_price': [100, 400, 900],
                'Category': ['Category1', 'Category2', 'Category3']
            })
            self.sheets = []
    
    # Create main window instance
    main_window = MainWindow(None)
    
    # Test master dataset creation
    master_df = main_window._create_unified_dataframe(MockFileMapping(), is_master=True)
    
    if master_df is not None:
        logger.info(f"✅ Master DataFrame created successfully: {len(master_df)} rows")
        logger.info(f"Columns: {list(master_df.columns)}")
        logger.info(f"Sample descriptions: {master_df['Description'].head(3).tolist()}")
        
        # Verify descriptions are not empty
        empty_descriptions = master_df['Description'].isna() | (master_df['Description'] == '')
        if empty_descriptions.any():
            logger.error(f"❌ Found {empty_descriptions.sum()} empty descriptions in master dataset")
            return False
        else:
            logger.info("✅ All master descriptions are non-empty")
    else:
        logger.error("❌ Failed to create master DataFrame")
        return False
    
    # Test comparison dataset creation
    comparison_df = main_window._create_unified_dataframe(MockFileMapping(), is_master=False)
    
    if comparison_df is not None:
        logger.info(f"✅ Comparison DataFrame created successfully: {len(comparison_df)} rows")
        logger.info(f"Columns: {list(comparison_df.columns)}")
        logger.info(f"Sample descriptions: {comparison_df['Description'].head(3).tolist()}")
        
        # Verify descriptions are not empty
        empty_descriptions = comparison_df['Description'].isna() | (comparison_df['Description'] == '')
        if empty_descriptions.any():
            logger.error(f"❌ Found {empty_descriptions.sum()} empty descriptions in comparison dataset")
            return False
        else:
            logger.info("✅ All comparison descriptions are non-empty")
    else:
        logger.error("❌ Failed to create comparison DataFrame")
        return False
    
    # Test column alignment
    master_cols = set(master_df.columns)
    comparison_cols = set(comparison_df.columns)
    
    if master_cols == comparison_cols:
        logger.info("✅ Column sets match between master and comparison datasets")
    else:
        missing_in_comparison = master_cols - comparison_cols
        missing_in_master = comparison_cols - master_cols
        logger.warning(f"Column mismatch - Missing in comparison: {missing_in_comparison}")
        logger.warning(f"Column mismatch - Missing in master: {missing_in_master}")
    
    return True

def test_comparison_processor():
    """Test the comparison processor with unified data"""
    logger.info("Testing comparison processor...")
    
    # Create test data
    master_df = pd.DataFrame({
        'Description': ['Item 1', 'Item 2', 'Item 3'],
        'code': ['CODE1', 'CODE2', 'CODE3'],
        'unit': ['EA', 'EA', 'EA'],
        'quantity': [1, 2, 3],
        'unit_price': [100, 200, 300],
        'total_price': [100, 400, 900]
    })
    
    comparison_df = pd.DataFrame({
        'Description': ['Item 1', 'Item 2', 'Item 4'],  # Item 4 is new
        'code': ['CODE1', 'CODE2', 'CODE4'],
        'unit': ['EA', 'EA', 'EA'],
        'quantity': [1, 2, 4],
        'unit_price': [100, 200, 400],
        'total_price': [100, 400, 1600]
    })
    
    # Create comparison processor
    processor = ComparisonProcessor()
    
    # Load datasets
    processor.load_master_dataset(master_df)
    processor.load_comparison_data(comparison_df)
    
    # Validate
    is_valid, message = processor.validate_comparison_data()
    if is_valid:
        logger.info("✅ Comparison data validation passed")
    else:
        logger.error(f"❌ Comparison data validation failed: {message}")
        return False
    
    # Process rows
    row_results = processor.process_comparison_rows()
    valid_rows = [r for r in row_results if r['is_valid']]
    logger.info(f"✅ Found {len(valid_rows)} valid rows out of {len(row_results)} total")
    
    # Process valid rows
    if valid_rows:
        instance_results = processor.process_valid_rows(offer_name="TestOffer")
        logger.info(f"✅ Processed {len(instance_results)} valid rows")
    else:
        logger.warning("⚠️ No valid rows to process")
    
    return True

if __name__ == "__main__":
    logger.info("Starting unified comparison tests...")
    
    # Test 1: Unified DataFrame creation
    if test_unified_dataframe_creation():
        logger.info("✅ Test 1 passed: Unified DataFrame creation")
    else:
        logger.error("❌ Test 1 failed: Unified DataFrame creation")
        sys.exit(1)
    
    # Test 2: Comparison processor
    if test_comparison_processor():
        logger.info("✅ Test 2 passed: Comparison processor")
    else:
        logger.error("❌ Test 2 failed: Comparison processor")
        sys.exit(1)
    
    logger.info("🎉 All tests passed! Unified comparison approach is working correctly.") 