#!/usr/bin/env python3
"""
Debug script to identify comparison mismatches
"""

import sys
import os
import logging
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(name)s: %(message)s')
logger = logging.getLogger(__name__)

def debug_description_matching(master_df, comparison_df):
    """Debug description matching between master and comparison datasets"""
    
    logger.info("=== DEBUGGING DESCRIPTION MATCHING ===")
    
    # Get all unique descriptions from both datasets
    master_descriptions = set(master_df['Description'].dropna().astype(str).tolist())
    comparison_descriptions = set(comparison_df['Description'].dropna().astype(str).tolist())
    
    logger.info(f"Master dataset has {len(master_descriptions)} unique descriptions")
    logger.info(f"Comparison dataset has {len(comparison_descriptions)} unique descriptions")
    
    # Find descriptions that exist in comparison but not in master
    only_in_comparison = comparison_descriptions - master_descriptions
    only_in_master = master_descriptions - comparison_descriptions
    
    if only_in_comparison:
        logger.warning(f"Found {len(only_in_comparison)} descriptions only in comparison dataset:")
        for i, desc in enumerate(list(only_in_comparison)[:5]):  # Show first 5
            logger.warning(f"  {i+1}. '{desc}'")
            if i >= 4:
                logger.warning(f"  ... and {len(only_in_comparison) - 5} more")
                break
    
    if only_in_master:
        logger.warning(f"Found {len(only_in_master)} descriptions only in master dataset:")
        for i, desc in enumerate(list(only_in_master)[:5]):  # Show first 5
            logger.warning(f"  {i+1}. '{desc}'")
            if i >= 4:
                logger.warning(f"  ... and {len(only_in_master) - 5} more")
                break
    
    # Check for whitespace issues
    logger.info("=== CHECKING FOR WHITESPACE ISSUES ===")
    
    for desc in list(only_in_comparison)[:3]:  # Check first 3
        logger.info(f"Checking description: '{desc}'")
        
        # Try different whitespace variations
        variations = [
            desc.strip(),
            desc.rstrip(),
            desc.lstrip(),
            desc.replace('\n', ' ').replace('\r', ' ').replace('\t', ' '),
            ' '.join(desc.split())
        ]
        
        for var in variations:
            if var in master_descriptions:
                logger.warning(f"  Found match with variation: '{var}'")
                break
        else:
            logger.info(f"  No whitespace variation found a match")
    
    # Check for case sensitivity issues
    logger.info("=== CHECKING FOR CASE SENSITIVITY ISSUES ===")
    
    master_descriptions_lower = {desc.lower() for desc in master_descriptions}
    comparison_descriptions_lower = {desc.lower() for desc in comparison_descriptions}
    
    only_in_comparison_lower = comparison_descriptions_lower - master_descriptions_lower
    
    if only_in_comparison_lower:
        logger.warning(f"Found {len(only_in_comparison_lower)} descriptions with case differences:")
        for i, desc_lower in enumerate(list(only_in_comparison_lower)[:3]):
            logger.warning(f"  {i+1}. '{desc_lower}' (lowercase)")
    
    # Check for special characters
    logger.info("=== CHECKING FOR SPECIAL CHARACTERS ===")
    
    for desc in list(only_in_comparison)[:3]:
        logger.info(f"Description: '{desc}'")
        logger.info(f"  Length: {len(desc)}")
        logger.info(f"  ASCII: {repr(desc)}")
        logger.info(f"  Unicode: {[ord(c) for c in desc[:10]]}")  # First 10 chars
    
    return len(only_in_comparison) == 0

def debug_column_structure(master_df, comparison_df):
    """Debug column structure differences"""
    
    logger.info("=== DEBUGGING COLUMN STRUCTURE ===")
    
    logger.info(f"Master columns: {list(master_df.columns)}")
    logger.info(f"Comparison columns: {list(comparison_df.columns)}")
    
    # Check if Description column exists in both
    if 'Description' not in master_df.columns:
        logger.error("❌ No 'Description' column in master dataset!")
        return False
    
    if 'Description' not in comparison_df.columns:
        logger.error("❌ No 'Description' column in comparison dataset!")
        return False
    
    logger.info("✅ Both datasets have 'Description' column")
    
    # Check data types
    logger.info(f"Master Description dtype: {master_df['Description'].dtype}")
    logger.info(f"Comparison Description dtype: {comparison_df['Description'].dtype}")
    
    # Check for null values
    master_nulls = master_df['Description'].isnull().sum()
    comparison_nulls = comparison_df['Description'].isnull().sum()
    
    logger.info(f"Master Description nulls: {master_nulls}")
    logger.info(f"Comparison Description nulls: {comparison_nulls}")
    
    return True

if __name__ == "__main__":
    logger.info("Debug script ready. Use this to analyze comparison mismatches.")
    logger.info("Call debug_description_matching(master_df, comparison_df) to analyze descriptions.")
    logger.info("Call debug_column_structure(master_df, comparison_df) to analyze columns.") 