#!/usr/bin/env python3
"""
Debug test for ComparisonProcessor row processing
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


def test_row_processing_debug():
    """Debug test for row processing logic"""
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
        
        # Debug each row
        for idx, row in comparison_df.iterrows():
            description = row['Description']
            logger.info(f"Row {idx}: Description='{description}'")
            
            # Debug manual invalidation matching
            for invalidation in processor.manual_invalidations:
                invalidation_desc = invalidation.split('|')[0] if '|' in invalidation else invalidation
                logger.info(f"  Checking invalidation: '{invalidation}' -> description: '{invalidation_desc}'")
                logger.info(f"  Description match: {description == invalidation_desc}")
            
            is_manually_invalid = any(description in invalidation.split('|')[0] for invalidation in processor.manual_invalidations)
            logger.info(f"  Is manually invalid: {is_manually_invalid}")
            
            if not is_manually_invalid:
                # Test ROW_VALIDITY directly
                from core.row_classifier import RowClassifier
                classifier = RowClassifier()
                column_mapping = {i: col for i, col in enumerate(comparison_df.columns)}
                row_values = [str(row[col]) if row[col] is not None else '' for col in comparison_df.columns]
                validity_result = classifier.ROW_VALIDITY(row_values, column_mapping)
                logger.info(f"  ROW_VALIDITY result: {validity_result}")
                logger.info(f"  Row values: {row_values}")
        
        # Test ROW_VALIDITY directly with the same data
        from core.row_classifier import RowClassifier
        from utils.config import ColumnType
        
        classifier = RowClassifier()
        
        # Test with the exact same data that's failing
        test_row = ['Steel Beams', '8', '220.0', '1760.0']
        test_mapping = {0: ColumnType.DESCRIPTION, 1: ColumnType.QUANTITY, 2: ColumnType.UNIT_PRICE, 3: ColumnType.TOTAL_PRICE}
        
        logger.info(f"Testing ROW_VALIDITY with: {test_row}")
        logger.info(f"Column mapping: {test_mapping}")
        
        result = classifier.ROW_VALIDITY(test_row, test_mapping)
        logger.info(f"ROW_VALIDITY result: {result}")
        
        # Test each condition individually
        for col_idx, col_type in test_mapping.items():
            if col_idx < len(test_row):
                cell_value = test_row[col_idx].strip() if test_row[col_idx] else ""
                logger.info(f"Column {col_type}: value='{cell_value}'")
                
                if col_type == ColumnType.DESCRIPTION:
                    has_desc = bool(cell_value)
                    logger.info(f"  Has description: {has_desc}")
                elif col_type == ColumnType.QUANTITY:
                    is_pos_num = classifier._is_positive_numeric(cell_value)
                    logger.info(f"  Is positive numeric: {is_pos_num}")
                elif col_type == ColumnType.UNIT_PRICE:
                    is_pos_num = classifier._is_positive_numeric(cell_value)
                    logger.info(f"  Is positive numeric: {is_pos_num}")
                elif col_type == ColumnType.TOTAL_PRICE:
                    is_pos_num = classifier._is_positive_numeric(cell_value)
                    logger.info(f"  Is positive numeric: {is_pos_num}")
        
        logger.info(f"Manual invalidations: {processor.manual_invalidations}")
        logger.info(f"Comparison data columns: {list(comparison_df.columns)}")
        
        # Process rows
        results = processor.process_comparison_rows()
        
        logger.info(f"Row results: {results}")
        
        # Check manual override results
        manual_override_results = [r for r in results if r['reason'] == 'MANUAL_OVERRIDE']
        logger.info(f"Manual override results: {manual_override_results}")
        assert len(manual_override_results) == 1, f"Expected 1 manual override, got {len(manual_override_results)}"
        
        logger.info("✅ Row processing logic: PASS")
        return True
        
    except Exception as e:
        logger.error(f"❌ Row processing test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    test_row_processing_debug() 