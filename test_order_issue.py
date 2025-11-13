#!/usr/bin/env python3
"""
Test script to verify the row order issue in comparison logic
"""

import pandas as pd
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_order_issue():
    """Test to verify the row order issue"""
    logger.info("=== TESTING ROW ORDER ISSUE ===")
    
    try:
        # Load the debug file
        debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_222545.xlsx"
        
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        logger.info(f"Master dataset shape: {master_df.shape}")
        logger.info(f"Comparison dataset shape: {comparison_df.shape}")
        
        # Find GEOTEXTILE rows in both datasets
        geotextile_desc = "GEO-SYNTHETICS: GEOTEXTILES"
        
        master_geotextile = master_df[master_df['Description'].str.contains('GEOTEXTILES', na=False)]
        comp_geotextile = comparison_df[comparison_df['Description'].str.contains('GEOTEXTILES', na=False)]
        
        logger.info(f"Master GEOTEXTILE rows: {len(master_geotextile)}")
        logger.info(f"Comparison GEOTEXTILE rows: {len(comp_geotextile)}")
        
        # Check the order and quantities
        logger.info("\n=== MASTER GEOTEXTILE ROWS ===")
        for idx, row in master_geotextile.iterrows():
            logger.info(f"Row {idx}: Quantity={row.get('quantity', 'N/A')}, Total={row.get('total_price', 'N/A')}")
        
        logger.info("\n=== COMPARISON GEOTEXTILE ROWS ===")
        for idx, row in comp_geotextile.iterrows():
            logger.info(f"Row {idx}: Quantity={row.get('quantity', 'N/A')}, Total={row.get('total_price', 'N/A')}")
        
        # Check the MERGE_ADD_Decision column
        if 'MERGE_ADD_Decision' in comparison_df.columns:
            logger.info("\n=== GEOTEXTILE MERGE/ADD DECISIONS ===")
            for idx, row in comp_geotextile.iterrows():
                decision = row.get('MERGE_ADD_Decision', 'N/A')
                logger.info(f"Row {idx}: {decision}")
        
        # The issue: The comparison dataset has the rows in a different order
        # than expected. The instance matching assumes first instance in master
        # matches first instance in comparison, but the order is wrong.
        
        logger.info("\n=== ANALYSIS ===")
        logger.info("The problem is that the comparison dataset rows are not in the same order")
        logger.info("as the original Excel file. The instance matching logic assumes:")
        logger.info("1st instance in master -> 1st instance in comparison")
        logger.info("2nd instance in master -> 2nd instance in comparison")
        logger.info("etc.")
        logger.info("But if the order is wrong, the matching fails.")
        
    except Exception as e:
        logger.error(f"Error in test: {e}")

if __name__ == "__main__":
    test_order_issue() 