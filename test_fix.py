#!/usr/bin/env python3
import pandas as pd
import logging
from core.comparison_engine import ComparisonProcessor

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_geotextile_fix():
    """Test if the GEOTEXTILE issue is fixed with normalized matching."""
    
    debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_215326.xlsx"
    
    try:
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        print("=== TESTING GEOTEXTILE FIX ===")
        
        # Find GEOTEXTILE rows
        geotextile_mask = comparison_df['Description'].str.contains('GEOTEXTILES', case=False, na=False)
        geotextile_rows = comparison_df[geotextile_mask]
        
        print(f"GEOTEXTILE rows in comparison: {len(geotextile_rows)}")
        
        # Test the new normalized matching logic
        description = "GEO-SYNTHETICS: GEOTEXTILES\nSupply and installation of geotextiles for reinforcement with weight per unit area according to the design in the whole width of the roads and pads, as shown on the drawings, including ancillary material for overlapping and proper installation."
        normalized_desc = ' '.join(description.split())
        
        print(f"Original description length: {len(description)}")
        print(f"Normalized description length: {len(normalized_desc)}")
        print(f"Normalized description: '{normalized_desc}'")
        
        # Test the new matching logic
        master_instances = master_df[
            master_df['Description'].str.lower().apply(lambda x: ' '.join(str(x).split())) == normalized_desc.lower()
        ]
        
        print(f"Master instances found with normalized matching: {len(master_instances)}")
        
        # Test the old exact matching logic
        old_master_instances = master_df[master_df['Description'] == description]
        print(f"Master instances found with exact matching: {len(old_master_instances)}")
        
        # Show the difference
        if len(master_instances) > len(old_master_instances):
            print("✅ FIX WORKING: Normalized matching found more instances than exact matching")
        else:
            print("❌ FIX NOT WORKING: No improvement in instance matching")
        
        # Test the full comparison processor
        print("\n=== TESTING FULL COMPARISON PROCESSOR ===")
        
        processor = ComparisonProcessor()
        processor.load_master_dataset(master_df)
        processor.load_comparison_data(comparison_df)
        
        # Process the comparison (need to validate rows first)
        processor.process_comparison_rows()
        results = processor.process_valid_rows()
        
        # Count MERGE vs ADD operations
        merge_count = len([r for r in results if r['type'] == 'MERGE'])
        add_count = len([r for r in results if r['type'] == 'ADD'])
        
        print(f"Total operations: {len(results)}")
        print(f"MERGE operations: {merge_count}")
        print(f"ADD operations: {add_count}")
        
        if add_count == 0:
            print("✅ SUCCESS: No ADD operations found - all rows should be MERGE")
        else:
            print(f"❌ ISSUE: Found {add_count} ADD operations when there should be none")
            
            # Check if any ADD operations are for GEOTEXTILE
            for result in results:
                if result['type'] == 'ADD':
                    comp_idx = result['comp_row_index']
                    comp_row = comparison_df.iloc[comp_idx]
                    if 'GEOTEXTILES' in str(comp_row.get('Description', '')):
                        print(f"  ADD operation for GEOTEXTILE at row {comp_idx + 1}")
        
    except Exception as e:
        print(f"Error testing fix: {e}")

if __name__ == "__main__":
    test_geotextile_fix() 