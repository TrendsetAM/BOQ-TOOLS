#!/usr/bin/env python3
"""
Test comparison using the UI workflow to reproduce the GEOTEXTILE ADD issue
"""

import pandas as pd
import logging
import os
from core.comparison_engine import ComparisonProcessor

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_ui_comparison_workflow():
    """Test comparison by simulating the UI workflow"""
    logger.info("=== TESTING UI COMPARISON WORKFLOW ===")
    
    try:
        # Use the specific BoQ file from examples folder
        boq_file = "examples/GRE.EEC.F.27.IT.P.18371.00.098.02 - PONTESTURA 9,69 MW_cBOQ PV_rev 9 giu (ENEMEK).xlsx"
        
        if not os.path.exists(boq_file):
            logger.error(f"BoQ file not found: {boq_file}")
            return None
            
        logger.info(f"Using BoQ file: {boq_file}")
        
        # Read the Excel file directly
        logger.info("Loading Excel file...")
        excel_data = pd.read_excel(boq_file, sheet_name=None)
        
        # Try to find a sheet that looks like a BoQ
        boq_sheet = None
        preferred_sheets = ['Civil Works', 'Solar-BOP', 'Miscellaneous', 'EMBOP_CI', 'EMBOP_SI']
        
        for sheet_name in preferred_sheets:
            if sheet_name in excel_data:
                boq_sheet = excel_data[sheet_name]
                logger.info(f"Using preferred sheet: {sheet_name}")
                break
        
        # If no preferred sheet found, use the first one with more than 10 rows
        if boq_sheet is None:
            for sheet_name, sheet_data in excel_data.items():
                if len(sheet_data) > 10:  # Assume BoQ sheets have more than 10 rows
                    boq_sheet = sheet_data
                    logger.info(f"Using fallback sheet: {sheet_name}")
                    break
        
        if boq_sheet is None:
            logger.error("No suitable BoQ sheet found")
            return None
            
        logger.info(f"BoQ sheet shape: {boq_sheet.shape}")
        
        # Create two identical datasets (simulating master and comparison)
        master_df = boq_sheet.copy()
        comparison_df = boq_sheet.copy()
        
        logger.info(f"Master dataset shape: {master_df.shape}")
        logger.info(f"Comparison dataset shape: {comparison_df.shape}")
        
        # Check if we have a Description column
        if 'Description' not in master_df.columns:
            logger.warning("No 'Description' column found. Looking for similar columns...")
            desc_columns = [col for col in master_df.columns if 'desc' in col.lower() or 'item' in col.lower()]
            if desc_columns:
                logger.info(f"Found potential description columns: {desc_columns}")
                # Use the first one as Description
                master_df['Description'] = master_df[desc_columns[0]]
                comparison_df['Description'] = comparison_df[desc_columns[0]]
            else:
                logger.error("No description column found")
                return None
        
        # Look for the GEOTEXTILE row
        geotextile_rows = master_df[master_df['Description'].str.contains('GEOTEXTILE', case=False, na=False)]
        logger.info(f"Found {len(geotextile_rows)} rows containing 'GEOTEXTILE'")
        
        for idx, row in geotextile_rows.iterrows():
            logger.info(f"GEOTEXTILE row {idx}: '{row['Description']}'")
        
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
    results = test_ui_comparison_workflow()
    if results:
        print(f"\nTest completed. Found {len(results)} total operations.")
    else:
        print("\nTest failed.") 