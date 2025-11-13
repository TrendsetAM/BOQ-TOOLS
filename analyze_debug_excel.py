#!/usr/bin/env python3
import pandas as pd
import numpy as np

def analyze_debug_excel():
    """Analyze the debug Excel file to understand ADD operations and row indexing issues."""
    
    debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_215326.xlsx"
    
    try:
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        print("=== DEBUG EXCEL ANALYSIS ===")
        print(f"Master dataset shape: {master_df.shape}")
        print(f"Comparison dataset shape: {comparison_df.shape}")
        
        print("\n=== MASTER DATASET COLUMNS ===")
        print(master_df.columns.tolist())
        
        print("\n=== COMPARISON DATASET COLUMNS ===")
        print(comparison_df.columns.tolist())
        
        # Check for MERGE_ADD_Decision column
        if 'MERGE_ADD_Decision' in comparison_df.columns:
            print("\n=== MERGE/ADD DECISIONS ===")
            decisions = comparison_df['MERGE_ADD_Decision'].value_counts()
            print(decisions)
            
            # Find ADD operations
            add_rows = comparison_df[comparison_df['MERGE_ADD_Decision'].str.contains('ADD', na=False)]
            if not add_rows.empty:
                print(f"\n=== ADD OPERATIONS ({len(add_rows)} found) ===")
                for idx, row in add_rows.iterrows():
                    print(f"Row {idx + 1}: {row.get('Description', 'N/A')} - {row.get('MERGE_ADD_Decision', 'N/A')}")
            
            # Find MERGE operations
            merge_rows = comparison_df[comparison_df['MERGE_ADD_Decision'].str.contains('MERGE', na=False)]
            if not merge_rows.empty:
                print(f"\n=== SAMPLE MERGE OPERATIONS (first 5) ===")
                for idx, row in merge_rows.head().iterrows():
                    print(f"Row {idx + 1}: {row.get('Description', 'N/A')} - {row.get('MERGE_ADD_Decision', 'N/A')}")
        
        # Check for row indexing issues
        print("\n=== ROW INDEXING ANALYSIS ===")
        if 'MERGE_ADD_Decision' in comparison_df.columns:
            # Look for patterns in the decision column
            sample_decisions = comparison_df['MERGE_ADD_Decision'].dropna().head(10)
            print("Sample decisions:")
            for i, decision in enumerate(sample_decisions):
                print(f"  {i+1}: {decision}")
        
        # Compare descriptions between master and comparison
        print("\n=== DESCRIPTION COMPARISON ===")
        if 'Description' in master_df.columns and 'Description' in comparison_df.columns:
            master_descriptions = set(master_df['Description'].dropna())
            comparison_descriptions = set(comparison_df['Description'].dropna())
            
            print(f"Master unique descriptions: {len(master_descriptions)}")
            print(f"Comparison unique descriptions: {len(comparison_descriptions)}")
            
            # Find descriptions that are in comparison but not in master (potential ADDs)
            only_in_comparison = comparison_descriptions - master_descriptions
            if only_in_comparison:
                print(f"\nDescriptions only in comparison (potential ADDs):")
                for desc in list(only_in_comparison)[:5]:  # Show first 5
                    print(f"  - {desc}")
            
            # Find descriptions that are in master but not in comparison (potential DELETEs)
            only_in_master = master_descriptions - comparison_descriptions
            if only_in_master:
                print(f"\nDescriptions only in master (potential DELETEs):")
                for desc in list(only_in_master)[:5]:  # Show first 5
                    print(f"  - {desc}")
        
    except Exception as e:
        print(f"Error analyzing debug Excel file: {e}")

if __name__ == "__main__":
    analyze_debug_excel() 