#!/usr/bin/env python3
"""
Test script for auto-categorization only workflow (no manual categorization needed)
"""

import sys
import os
from pathlib import Path
import pandas as pd

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.auto_categorizer import auto_categorize_dataset, collect_descriptions_for_manual_review
from core.manual_categorizer import execute_row_categorization
from core.category_dictionary import CategoryDictionary


def test_auto_categorization_only():
    """Test the workflow when all rows are auto-categorized"""
    print("=== Testing Auto-Categorization Only Workflow ===\n")
    
    # 1. Load category dictionary
    print("1. Loading category dictionary...")
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    initial_dict_size = len(category_dict.mappings)
    print(f"   Initial dictionary size: {initial_dict_size}")
    
    # 2. Create sample data with descriptions that should be auto-categorized
    print("\n2. Creating sample data with known descriptions...")
    sample_data = {
        'Description': [
            'cost of bank guarantee (down payment, execution, warranty) and insurance',  # From dictionary
            'standard safety costs',  # From dictionary
            'detail design (civil design, electrical design, cable design etc as per dpp) & as-built design',  # From dictionary
            'project and site management (project management, h&seq, env. coord, activities planning, material control, contract administration….)',  # From dictionary
            'execution of the environmental works, environmental mitigation and restoration works in accordance to the technical specification.'  # From dictionary
        ],
        'Quantity': [10, 100, 50, 5, 2],
        'Unit_Price': [100, 25, 200, 150, 75],
        'Total_Price': [1000, 2500, 10000, 750, 150]
    }
    df = pd.DataFrame(sample_data)
    print(f"   Created DataFrame with {len(df)} rows")
    
    # 3. Test the workflow
    print("\n3. Testing workflow (should auto-categorize all rows)...")
    result = execute_row_categorization(df)
    
    if result['error']:
        print(f"   ❌ Error: {result['error']}")
        return
    
    print("   ✅ Auto-categorization completed")
    
    # 4. Check if manual categorization was needed
    manual_needed = result.get('summary', {}).get('manual_categorization_needed', True)
    all_auto = result.get('summary', {}).get('all_auto_categorized', False)
    manual_excel_path = result.get('manual_excel_path')
    
    print(f"   Manual categorization needed: {manual_needed}")
    print(f"   All auto-categorized: {all_auto}")
    print(f"   Manual Excel path: {manual_excel_path}")
    
    # 5. Check statistics
    stats = result.get('all_stats', {})
    auto_stats = stats.get('auto_stats', {})
    total_rows = auto_stats.get('total_rows', 0)
    matched_rows = auto_stats.get('matched_rows', 0)
    unmatched_rows = total_rows - matched_rows
    
    print(f"\n4. Statistics:")
    print(f"   Total rows: {total_rows}")
    print(f"   Matched rows: {matched_rows}")
    print(f"   Unmatched rows: {unmatched_rows}")
    print(f"   Match rate: {matched_rows/total_rows:.1%}" if total_rows > 0 else "   Match rate: N/A")
    
    # 6. Verify the result
    if manual_needed == False and all_auto == True and manual_excel_path is None:
        print("\n✅ SUCCESS: All rows were auto-categorized!")
        print("✅ Manual categorization was correctly skipped")
        print("✅ No Excel file was generated")
    else:
        print("\n❌ FAILURE: Manual categorization was not skipped as expected")
        print(f"   Expected: manual_needed=False, all_auto=True, manual_excel_path=None")
        print(f"   Actual: manual_needed={manual_needed}, all_auto={all_auto}, manual_excel_path={manual_excel_path}")
    
    print("\n=== Test Complete ===")


if __name__ == "__main__":
    test_auto_categorization_only() 