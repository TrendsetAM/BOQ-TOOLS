#!/usr/bin/env python3
import pandas as pd

def analyze_geotextile_issue():
    """Analyze the specific GEOTEXTILE row that's being marked as ADD."""
    
    debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_215326.xlsx"
    
    try:
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        print("=== GEOTEXTILE ISSUE ANALYSIS ===")
        
        # Find the GEOTEXTILE row in comparison dataset
        geotextile_mask = comparison_df['Description'].str.contains('GEOTEXTILES', case=False, na=False)
        geotextile_rows = comparison_df[geotextile_mask]
        
        print(f"GEOTEXTILE rows in comparison dataset: {len(geotextile_rows)}")
        for idx, row in geotextile_rows.iterrows():
            print(f"Row {idx + 1}: {row['Description'][:100]}...")
            print(f"Decision: {row.get('MERGE_ADD_Decision', 'N/A')}")
            print(f"Code: {row.get('code', 'N/A')}")
            print(f"Category: {row.get('Category', 'N/A')}")
            print("---")
        
        # Find GEOTEXTILE rows in master dataset
        master_geotextile_mask = master_df['Description'].str.contains('GEOTEXTILES', case=False, na=False)
        master_geotextile_rows = master_df[master_geotextile_mask]
        
        print(f"GEOTEXTILE rows in master dataset: {len(master_geotextile_rows)}")
        for idx, row in master_geotextile_rows.iterrows():
            print(f"Row {idx + 1}: {row['Description'][:100]}...")
            print(f"Code: {row.get('code', 'N/A')}")
            print(f"Category: {row.get('Category', 'N/A')}")
            print("---")
        
        # Compare the specific GEOTEXTILE descriptions
        print("=== DETAILED COMPARISON ===")
        if len(geotextile_rows) > 0 and len(master_geotextile_rows) > 0:
            comp_desc = geotextile_rows.iloc[0]['Description']
            master_desc = master_geotextile_rows.iloc[0]['Description']
            
            print("Comparison Description:")
            print(comp_desc)
            print("\nMaster Description:")
            print(master_desc)
            print(f"\nDescriptions are identical: {comp_desc == master_desc}")
            
            # Check other fields
            comp_row = geotextile_rows.iloc[0]
            master_row = master_geotextile_rows.iloc[0]
            
            print(f"\nCode comparison: {comp_row['code']} vs {master_row['code']}")
            print(f"Category comparison: {comp_row['Category']} vs {master_row['Category']}")
            print(f"Unit comparison: {comp_row['unit']} vs {master_row['unit']}")
            print(f"Quantity comparison: {comp_row['quantity']} vs {master_row['quantity']}")
        
        # Check for duplicate descriptions in master
        print("\n=== DUPLICATE ANALYSIS ===")
        master_desc_counts = master_df['Description'].value_counts()
        geotextile_count = master_desc_counts.get(comp_desc, 0) if len(geotextile_rows) > 0 else 0
        print(f"GEOTEXTILE description appears {geotextile_count} times in master dataset")
        
        # Check for duplicate descriptions in comparison
        comp_desc_counts = comparison_df['Description'].value_counts()
        geotextile_comp_count = comp_desc_counts.get(comp_desc, 0) if len(geotextile_rows) > 0 else 0
        print(f"GEOTEXTILE description appears {geotextile_comp_count} times in comparison dataset")
        
    except Exception as e:
        print(f"Error analyzing GEOTEXTILE issue: {e}")

if __name__ == "__main__":
    analyze_geotextile_issue() 