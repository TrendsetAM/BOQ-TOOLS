#!/usr/bin/env python3
"""
Simple debug test for MERGE/ADD logic
"""

import pandas as pd

def test_simple():
    print("=== Simple Test ===")
    
    # Simple test case
    master_df = pd.DataFrame({
        'Description': ['Item A', 'Item B'],
        'Quantity': [10, 5]
    })
    
    comparison_df = pd.DataFrame({
        'Description': ['Item A', 'Item B', 'Item C'],
        'Quantity': [12, 8, 20]
    })
    
    def determine_merge_add_decision(comparison_df, master_df):
        decisions = []
        
        for idx, comp_row in comparison_df.iterrows():
            description = str(comp_row.get('Description', '')).strip()
            print(f"Processing: '{description}'")
            
            # Get instances
            comp_instances = comparison_df[comparison_df['Description'] == description]
            master_instances = master_df[master_df['Description'] == description]
            
            print(f"  Comp instances: {len(comp_instances)}")
            print(f"  Master instances: {len(master_instances)}")
            
            # Find instance number
            comp_instance_number = 0
            for comp_idx, comp_instance_row in comp_instances.iterrows():
                if comp_idx == idx:
                    break
                comp_instance_number += 1
            
            print(f"  Instance number: {comp_instance_number}")
            
            # Decision
            if comp_instance_number < len(master_instances):
                master_idx = master_instances.index[comp_instance_number]
                decision = f'MERGE (Master Row {master_idx + 1})'
            else:
                decision = 'ADD (New Row)'
            
            print(f"  Decision: {decision}")
            decisions.append(decision)
        
        return decisions
    
    decisions = determine_merge_add_decision(comparison_df, master_df)
    
    print("\nFinal Results:")
    for i, (desc, decision) in enumerate(zip(comparison_df['Description'], decisions)):
        print(f"  {i+1}. '{desc}' → {decision}")

if __name__ == "__main__":
    test_simple() 