#!/usr/bin/env python3
"""
Debug test for comparison logic - specifically for the row 195 ADD issue
"""

import pandas as pd
import logging
from core.comparison_engine import ComparisonProcessor

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_same_file_comparison():
    """Test comparison with the same file loaded twice"""
    logger.info("=== STARTING SAME FILE COMPARISON TEST ===")
    
    try:
        # Load your debug file
        debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_211350.xlsx"
        
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        logger.info(f"Loaded master dataset: {master_df.shape}")
        logger.info(f"Loaded comparison dataset: {comparison_df.shape}")
        
        # Create processor
        processor = ComparisonProcessor()
        
        # Load datasets
        processor.load_master_dataset(master_df)
        processor.load_comparison_data(comparison_df)
        
        # Process rows
        logger.info("Processing comparison rows...")
        row_results = processor.process_comparison_rows()
        
        # Filter valid rows
        valid_rows = [r for r in row_results if r['is_valid']]
        logger.info(f"Valid rows: {len(valid_rows)}")
        
        # Process valid rows
        logger.info("Processing valid rows with MERGE/ADD logic...")
        results = processor.process_valid_rows(offer_name="TestOffer")
        
        # Analyze results
        merge_ops = [r for r in results if r['type'] == 'MERGE']
        add_ops = [r for r in results if r['type'] == 'ADD']
        
        logger.info(f"=== FINAL RESULTS ===")
        logger.info(f"MERGE operations: {len(merge_ops)}")
        logger.info(f"ADD operations: {len(add_ops)}")
        
        if len(add_ops) > 0:
            logger.warning(f"FOUND {len(add_ops)} ADD OPERATIONS!")
            for add_op in add_ops:
                logger.warning(f"ADD operation for comparison row {add_op['comp_row_index']}")
                
                # Get the row data
                comp_row = comparison_df.iloc[add_op['comp_row_index']]
                logger.warning(f"ADD row description: '{comp_row.get('Description', 'N/A')}'")
        
        return results
        
    except Exception as e:
        logger.error(f"Error in test: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    results = test_same_file_comparison()
    if results:
        print(f"\nTest completed. Found {len(results)} total operations.")
    else:
        print("\nTest failed.") 