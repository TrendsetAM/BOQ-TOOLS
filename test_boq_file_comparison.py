#!/usr/bin/env python3
"""
Test comparison using the actual Excel BoQ file to reproduce the GEOTEXTILE ADD issue
"""

import pandas as pd
import logging
from core.comparison_engine import ComparisonProcessor
from core.boq_processor import BoQProcessor
from core.file_processor import FileProcessor

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_boq_file_comparison():
    """Test comparison by loading the actual Excel BoQ file twice"""
    logger.info("=== TESTING BOQ FILE COMPARISON ===")
    
    try:
        # You'll need to specify the path to your actual Excel BoQ file
        boq_file_path = input("Enter the path to your Excel BoQ file: ").strip()
        
        if not boq_file_path:
            logger.error("No file path provided")
            return None
            
        logger.info(f"Loading BoQ file: {boq_file_path}")
        
        # Create file processor and BoQ processor
        file_processor = FileProcessor()
        boq_processor = BoQProcessor()
        
        # Load the file twice to simulate master and comparison
        logger.info("Loading as master dataset...")
        master_file_mapping = file_processor.process_file(boq_file_path)
        master_df = boq_processor.process_boq_file(master_file_mapping)
        
        logger.info("Loading as comparison dataset...")
        comparison_file_mapping = file_processor.process_file(boq_file_path)
        comparison_df = boq_processor.process_boq_file(comparison_file_mapping)
        
        logger.info(f"Master dataset shape: {master_df.shape}")
        logger.info(f"Comparison dataset shape: {comparison_df.shape}")
        
        # Create comparison processor
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
                
                # Check if it's the GEOTEXTILE row
                if 'GEOTEXTILE' in str(comp_row.get('Description', '')):
                    logger.error(f"FOUND THE GEOTEXTILE ADD ISSUE! Row {add_op['comp_row_index']}")
                    logger.error(f"Description: '{comp_row.get('Description', 'N/A')}'")
        
        return results
        
    except Exception as e:
        logger.error(f"Error in test: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    results = test_boq_file_comparison()
    if results:
        print(f"\nTest completed. Found {len(results)} total operations.")
    else:
        print("\nTest failed.") 