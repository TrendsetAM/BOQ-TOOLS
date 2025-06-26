#!/usr/bin/env python3
"""
Test script to verify debug prints in auto_categorize_dataset
"""

import sys
from pathlib import Path
import pandas as pd

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary
from core.auto_categorizer import auto_categorize_dataset

def test_debug_prints():
    """Test that debug prints are working in auto_categorize_dataset"""
    print("=== Testing Debug Prints in Auto Categorization ===\n")
    
    # Create a simple test DataFrame with descriptions that should match the dictionary
    test_data = {
        'Description': [
            'cost of bank guarantee (down payment, execution, warranty) and insurance',  # Should be exact match
            'standard safety costs',  # Should be exact match
            'site camp construction (including offices, main storage area, etc)',  # Should be exact match
            'cost of bank guarantee down payment execution warranty and insurance',  # Should be fuzzy match (missing parentheses)
            'standard safety cost',  # Should be fuzzy match (singular)
            'site camp construction including offices main storage area',  # Should be fuzzy match (missing parentheses)
            'unknown material xyz',  # Should be unmatched
            'mystery component abc',  # Should be unmatched
        ]
    }
    
    df = pd.DataFrame(test_data)
    print(f"Test DataFrame with {len(df)} rows:")
    print(df)
    print()
    
    # Initialize category dictionary
    print("Initializing category dictionary...")
    category_dict = CategoryDictionary()
    print(f"Loaded {len(category_dict.mappings)} mappings")
    print()
    
    # Run auto categorization
    print("Running auto categorization (debug prints should appear below):")
    print("-" * 60)
    
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.8
    )
    
    print("-" * 60)
    print("Auto categorization completed!")
    print()
    
    # Show results
    print("Results:")
    print(f"  Total rows: {result.total_rows}")
    print(f"  Matched rows: {result.matched_rows}")
    print(f"  Unmatched rows: {result.unmatched_rows}")
    print(f"  Match rate: {result.match_rate:.1%}")
    print()
    
    print("Categorized DataFrame:")
    print(result.dataframe[['Description', 'Category']])
    print()
    
    print("Match type breakdown:")
    for match_type, count in result.match_statistics['match_types'].items():
        print(f"  {match_type}: {count}")
    
    print("\n✅ Debug print test completed!")

if __name__ == "__main__":
    test_debug_prints() 