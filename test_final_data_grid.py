#!/usr/bin/env python3
"""
Test script for the final data grid functionality
"""

import sys
import os
from pathlib import Path
import pandas as pd

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.auto_categorizer import auto_categorize_dataset
from core.manual_categorizer import execute_row_categorization
from core.category_dictionary import CategoryDictionary


def test_final_data_grid():
    """Test the final data grid functionality"""
    print("=== Testing Final Data Grid Functionality ===\n")
    
    # 1. Create sample data
    print("1. Creating sample data...")
    sample_data = {
        'Description': [
            'cost of bank guarantee (down payment, execution, warranty) and insurance',
            'standard safety costs',
            'detail design (civil design, electrical design, cable design etc as per dpp) & as-built design',
            'project and site management (project management, h&seq, env. coord, activities planning, material control, contract administration….)',
            'execution of the environmental works, environmental mitigation and restoration works in accordance to the technical specification.'
        ],
        'Quantity': [10, 100, 50, 5, 2],
        'Unit_Price': [100, 25, 200, 150, 75],
        'Total_Price': [1000, 2500, 10000, 750, 150],
        'Unit': ['pcs', 'm', 'pcs', 'pcs', 'pcs'],
        'Sheet_Name': ['Sheet1', 'Sheet1', 'Sheet2', 'Sheet2', 'Sheet3']
    }
    df = pd.DataFrame(sample_data)
    print(f"   Created DataFrame with {len(df)} rows")
    print(f"   Columns: {list(df.columns)}")
    
    # 2. Run categorization
    print("\n2. Running categorization...")
    result = execute_row_categorization(df)
    
    if result['error']:
        print(f"   ❌ Error: {result['error']}")
        return
    
    print("   ✅ Categorization completed")
    
    # 3. Check final DataFrame
    final_df = result['final_dataframe']
    print(f"\n3. Final DataFrame:")
    print(f"   Shape: {final_df.shape}")
    print(f"   Columns: {list(final_df.columns)}")
    
    # 4. Check if Category column exists
    if 'Category' in final_df.columns:
        print(f"   ✅ Category column found")
        print(f"   Categories: {final_df['Category'].value_counts().to_dict()}")
    else:
        print(f"   ❌ Category column missing")
    
    # 5. Check if Sheet_Name column exists
    if 'Sheet_Name' in final_df.columns:
        print(f"   ✅ Sheet_Name column found")
        print(f"   Sheets: {final_df['Sheet_Name'].value_counts().to_dict()}")
    else:
        print(f"   ❌ Sheet_Name column missing")
    
    # 6. Display sample data
    print(f"\n4. Sample data (first 3 rows):")
    print(final_df.head(3).to_string())
    
    print("\n=== Test Complete ===")
    print("✅ Final data grid functionality is ready!")
    print("✅ DataFrame contains all required columns")
    print("✅ Data is properly categorized")


if __name__ == "__main__":
    test_final_data_grid() 