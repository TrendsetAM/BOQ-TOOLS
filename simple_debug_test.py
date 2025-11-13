#!/usr/bin/env python3
"""
Simple debug test for the row 195 ADD issue
"""

import pandas as pd
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def simple_test():
    """Simple test to check the debug file and row 195"""
    logger.info("=== SIMPLE DEBUG TEST ===")
    
    try:
        # Load the debug file
        debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_211350.xlsx"
        
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        logger.info(f"Master dataset shape: {master_df.shape}")
        logger.info(f"Comparison dataset shape: {comparison_df.shape}")
        
        # Check if row 195 exists in both datasets
        if len(master_df) > 195:
            master_row_195 = master_df.iloc[194]  # 0-indexed
            logger.info(f"Master row 195 description: '{master_row_195.get('Description', 'N/A')}'")
        else:
            logger.warning("Master dataset doesn't have row 195")
            
        if len(comparison_df) > 195:
            comp_row_195 = comparison_df.iloc[194]  # 0-indexed
            logger.info(f"Comparison row 195 description: '{comp_row_195.get('Description', 'N/A')}'")
        else:
            logger.warning("Comparison dataset doesn't have row 195")
        
        # Check for duplicate descriptions
        master_descriptions = master_df['Description'].value_counts()
        comp_descriptions = comparison_df['Description'].value_counts()
        
        logger.info(f"Master unique descriptions: {len(master_descriptions)}")
        logger.info(f"Comparison unique descriptions: {len(comp_descriptions)}")
        
        # Find descriptions that appear more than once
        master_duplicates = master_descriptions[master_descriptions > 1]
        comp_duplicates = comp_descriptions[comp_descriptions > 1]
        
        logger.info(f"Master descriptions with duplicates: {len(master_duplicates)}")
        logger.info(f"Comparison descriptions with duplicates: {len(comp_duplicates)}")
        
        if len(master_duplicates) > 0:
            logger.info("Master duplicates:")
            for desc, count in master_duplicates.head(5).items():
                logger.info(f"  '{desc}': {count} instances")
                
        if len(comp_duplicates) > 0:
            logger.info("Comparison duplicates:")
            for desc, count in comp_duplicates.head(5).items():
                logger.info(f"  '{desc}': {count} instances")
        
        return True
        
    except Exception as e:
        logger.error(f"Error in simple test: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = simple_test()
    if success:
        print("\nSimple test completed successfully.")
    else:
        print("\nSimple test failed.") 