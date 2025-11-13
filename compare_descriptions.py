#!/usr/bin/env python3
import pandas as pd

def compare_descriptions():
    """Compare GEOTEXTILE description with other multi-instance descriptions."""
    
    debug_file = "DEBUG_Datasets_Before_Merge_222_20250728_215326.xlsx"
    
    try:
        # Read both sheets
        master_df = pd.read_excel(debug_file, sheet_name='Master_Dataset')
        comparison_df = pd.read_excel(debug_file, sheet_name='Comparison_Dataset')
        
        print("=== DESCRIPTION COMPARISON ANALYSIS ===")
        
        # Find descriptions with multiple instances
        master_desc_counts = master_df['Description'].value_counts()
        comparison_desc_counts = comparison_df['Description'].value_counts()
        
        # Find descriptions that appear multiple times
        multi_instance_descriptions = master_desc_counts[master_desc_counts > 1]
        
        print(f"Descriptions with multiple instances in master: {len(multi_instance_descriptions)}")
        print("Top 10 multi-instance descriptions:")
        for desc, count in multi_instance_descriptions.head(10).items():
            print(f"  '{desc[:50]}...' - {count} instances")
        
        # Focus on GEOTEXTILE
        geotextile_desc = None
        for desc in master_df['Description'].unique():
            if 'GEOTEXTILES' in str(desc):
                geotextile_desc = desc
                break
        
        if geotextile_desc:
            print(f"\n=== GEOTEXTILE DESCRIPTION ANALYSIS ===")
            print(f"GEOTEXTILE description: '{geotextile_desc}'")
            print(f"Length: {len(geotextile_desc)}")
            print(f"Type: {type(geotextile_desc)}")
            
            # Check for hidden characters
            print(f"ASCII representation: {repr(geotextile_desc)}")
            
            # Check if it exists in comparison with exact match
            exact_match_in_comparison = comparison_df[comparison_df['Description'] == geotextile_desc]
            print(f"Exact matches in comparison: {len(exact_match_in_comparison)}")
            
            # Check case-insensitive match
            case_insensitive_match = comparison_df[
                comparison_df['Description'].str.lower() == geotextile_desc.lower()
            ]
            print(f"Case-insensitive matches in comparison: {len(case_insensitive_match)}")
            
            # Check normalized whitespace match
            normalized_desc = ' '.join(geotextile_desc.split())
            normalized_match = comparison_df[
                comparison_df['Description'].apply(lambda x: ' '.join(str(x).split())) == normalized_desc
            ]
            print(f"Normalized whitespace matches in comparison: {len(normalized_match)}")
            
            # Compare with a working multi-instance description
            print(f"\n=== COMPARING WITH WORKING DESCRIPTIONS ===")
            working_descriptions = []
            for desc, count in multi_instance_descriptions.head(5).items():
                if 'GEOTEXTILES' not in str(desc):  # Exclude GEOTEXTILE
                    working_descriptions.append(desc)
                    break
            
            if working_descriptions:
                working_desc = working_descriptions[0]
                print(f"Working description: '{working_desc}'")
                print(f"Length: {len(working_desc)}")
                print(f"Type: {type(working_desc)}")
                print(f"ASCII representation: {repr(working_desc)}")
                
                # Check if working description has exact matches
                working_exact_match = comparison_df[comparison_df['Description'] == working_desc]
                print(f"Working description exact matches in comparison: {len(working_exact_match)}")
                
                # Check if working description has case-insensitive matches
                working_case_insensitive = comparison_df[
                    comparison_df['Description'].str.lower() == working_desc.lower()
                ]
                print(f"Working description case-insensitive matches: {len(working_case_insensitive)}")
        
        # Check for any descriptions that have different counts between master and comparison
        print(f"\n=== COUNT MISMATCH ANALYSIS ===")
        for desc in master_desc_counts.index:
            if master_desc_counts[desc] > 1:  # Only check multi-instance descriptions
                master_count = master_desc_counts[desc]
                comparison_count = comparison_desc_counts.get(desc, 0)
                
                if master_count != comparison_count:
                    print(f"MISMATCH: '{desc[:50]}...'")
                    print(f"  Master: {master_count}, Comparison: {comparison_count}")
                    
                    # Check if it's GEOTEXTILE
                    if 'GEOTEXTILES' in str(desc):
                        print(f"  *** THIS IS THE GEOTEXTILE DESCRIPTION ***")
        
    except Exception as e:
        print(f"Error comparing descriptions: {e}")

if __name__ == "__main__":
    compare_descriptions() 