#!/usr/bin/env python3
import pandas as pd

def test_geotextile_instances():
    """Test to verify GEOTEXTILE instance counts and logic."""
    
    debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_215326.xlsx"
    
    try:
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        print("=== GEOTEXTILE INSTANCE ANALYSIS ===")
        
        # Find GEOTEXTILE rows in both datasets
        geotextile_desc = "GEO-SYNTHETICS: GEOTEXTILES"
        
        master_geotextile = master_df[master_df['Description'].str.contains('GEOTEXTILES', case=False, na=False)]
        comparison_geotextile = comparison_df[comparison_df['Description'].str.contains('GEOTEXTILES', case=False, na=False)]
        
        print(f"Master GEOTEXTILE instances: {len(master_geotextile)}")
        print(f"Comparison GEOTEXTILE instances: {len(comparison_geotextile)}")
        
        print("\n=== MASTER GEOTEXTILE ROWS ===")
        for idx, row in master_geotextile.iterrows():
            print(f"Row {idx + 1}: {row['Description'][:100]}...")
            print(f"  Code: {row.get('code', 'N/A')}")
            print(f"  Category: {row.get('Category', 'N/A')}")
        
        print("\n=== COMPARISON GEOTEXTILE ROWS ===")
        for idx, row in comparison_geotextile.iterrows():
            print(f"Row {idx + 1}: {row['Description'][:100]}...")
            print(f"  Decision: {row.get('MERGE_ADD_Decision', 'N/A')}")
            print(f"  Code: {row.get('code', 'N/A')}")
            print(f"  Category: {row.get('Category', 'N/A')}")
        
        # Test the exact logic from determine_merge_add_decision
        print("\n=== TESTING DECISION LOGIC ===")
        description_instance_counts = {}
        
        for idx, comp_row in comparison_geotextile.iterrows():
            description = str(comp_row.get('Description', '')).strip()
            
            # Get master instances
            master_instances = master_df[master_df['Description'] == description]
            
            # Initialize instance count for this description if not seen before
            if description not in description_instance_counts:
                description_instance_counts[description] = 0
            
            # Get the current instance number for this description
            comp_instance_number = description_instance_counts[description]
            description_instance_counts[description] += 1
            
            print(f"Instance {comp_instance_number + 1}:")
            print(f"  Master instances found: {len(master_instances)}")
            print(f"  Decision condition: {comp_instance_number} < {len(master_instances)} = {comp_instance_number < len(master_instances)}")
            
            if comp_instance_number < len(master_instances):
                master_idx = master_instances.index[comp_instance_number]
                print(f"  DECISION: MERGE Instance {comp_instance_number + 1} (Master Row {master_idx + 1})")
            else:
                print(f"  DECISION: ADD Instance {comp_instance_number + 1} (New Row)")
        
        # Check if descriptions are exactly identical
        print("\n=== DESCRIPTION COMPARISON ===")
        if len(master_geotextile) > 0 and len(comparison_geotextile) > 0:
            master_desc = master_geotextile.iloc[0]['Description']
            comparison_desc = comparison_geotextile.iloc[0]['Description']
            
            print(f"Master description: '{master_desc}'")
            print(f"Comparison description: '{comparison_desc}'")
            print(f"Descriptions are identical: {master_desc == comparison_desc}")
            print(f"Descriptions are equal (case-insensitive): {master_desc.lower() == comparison_desc.lower()}")
            
            # Check for whitespace differences
            master_normalized = ' '.join(master_desc.split())
            comparison_normalized = ' '.join(comparison_desc.split())
            print(f"Normalized descriptions are identical: {master_normalized == comparison_normalized}")
        
    except Exception as e:
        print(f"Error analyzing GEOTEXTILE instances: {e}")

if __name__ == "__main__":
    test_geotextile_instances() 