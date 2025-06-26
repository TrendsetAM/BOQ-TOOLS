#!/usr/bin/env python3
"""
Collect Unmatched Descriptions Demo
Demonstrates the collect_unmatched_descriptions function
"""

import sys
import pandas as pd
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary
from core.auto_categorizer import auto_categorize_dataset, collect_unmatched_descriptions
from utils.logger import setup_logging


def create_test_dataframe():
    """Create a test DataFrame with some unmatched descriptions"""
    test_data = {
        'Description': [
            'Concrete foundation work',      # Should match
            'Steel reinforcement supply',    # Should match
            'Electrical installation',       # Should match
            'Unknown material type',         # Won't match
            'Custom electrical work',        # Won't match
            'Site camp construction',        # Should match
            'Unknown material type',         # Won't match (duplicate)
            'Custom electrical work',        # Won't match (duplicate)
            'New unknown item',              # Won't match
            'Another unknown item',          # Won't match
            'Concrete foundation work',      # Should match (duplicate)
            'Steel reinforcement supply',    # Should match (duplicate)
            'MV cable installation',         # Should match
            'Transformer installation',      # Should match
            'Road construction',             # Should match
            'Foundation excavation',         # Should match
            'Testing and certification',     # Should match
            'Design and engineering',        # Should match
            'Plumbing installation',         # Should match
            'Carpentry work'                 # Should match
        ],
        'Quantity': [100, 50, 200, 25, 150, 1, 25, 150, 75, 30, 100, 50, 500, 5, 200, 300, 1, 1, 100, 80],
        'Unit': ['m3', 'tons', 'm', 'units', 'm', 'unit', 'units', 'm', 'units', 'units', 'm3', 'tons', 'm', 'units', 'm2', 'm3', 'unit', 'unit', 'm', 'm2'],
        'Sheet_Name': ['BOQ_Sheet1', 'BOQ_Sheet1', 'BOQ_Sheet1', 'BOQ_Sheet1', 'BOQ_Sheet1', 
                      'BOQ_Sheet2', 'BOQ_Sheet2', 'BOQ_Sheet2', 'BOQ_Sheet2', 'BOQ_Sheet2',
                      'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3',
                      'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3', 'BOQ_Sheet3']
    }
    
    return pd.DataFrame(test_data)


def demo_collect_unmatched_descriptions():
    """Demonstrate collecting unmatched descriptions"""
    print("=== Collect Unmatched Descriptions Demo ===\n")
    
    # Initialize category dictionary
    category_dict = CategoryDictionary()
    
    # Create test DataFrame
    df = create_test_dataframe()
    print(f"Test DataFrame with {len(df)} rows:")
    print(df[['Description', 'Sheet_Name']].head(10))
    print()
    
    # Perform auto categorization first
    print("Step 1: Auto-categorizing the DataFrame...")
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.7
    )
    
    print(f"Categorization results:")
    print(f"  Total rows: {result.total_rows}")
    print(f"  Matched rows: {result.matched_rows} ({result.match_rate:.1%})")
    print(f"  Unmatched rows: {result.unmatched_rows}")
    print()
    
    # Show categorized DataFrame
    print("Categorized DataFrame:")
    categorized_df = result.dataframe[['Description', 'Category', 'Sheet_Name']]
    print(categorized_df.head(10))
    print()
    
    # Step 2: Collect unmatched descriptions
    print("Step 2: Collecting unmatched descriptions...")
    unmatched_descriptions = collect_unmatched_descriptions(
        dataframe=result.dataframe,
        category_column='Category',
        description_column='Description',
        sheet_name_column='Sheet_Name'
    )
    
    print(f"\nUnmatched Descriptions Analysis:")
    print(f"  Total unique unmatched descriptions: {len(unmatched_descriptions)}")
    print(f"  Total frequency: {sum(desc.frequency for desc in unmatched_descriptions)}")
    print()
    
    # Show unmatched descriptions with metadata
    print("Unmatched Descriptions with Metadata:")
    print("-" * 80)
    for i, desc in enumerate(unmatched_descriptions):
        print(f"{i+1}. Description: '{desc.description}'")
        print(f"   Source Sheet: {desc.source_sheet_name}")
        print(f"   Frequency: {desc.frequency}")
        print(f"   Sample Rows: {desc.sample_rows}")
        print(f"   Original Index: {desc.original_index}")
        print()
    
    # Show frequency statistics
    if unmatched_descriptions:
        frequencies = [desc.frequency for desc in unmatched_descriptions]
        print(f"Frequency Statistics:")
        print(f"  Max frequency: {max(frequencies)}")
        print(f"  Min frequency: {min(frequencies)}")
        print(f"  Average frequency: {sum(frequencies) / len(frequencies):.2f}")
        print()
        
        # Show most frequent unmatched descriptions
        print("Most Frequent Unmatched Descriptions:")
        for i, desc in enumerate(unmatched_descriptions[:3]):
            print(f"  {i+1}. '{desc.description}' (frequency: {desc.frequency})")
        print()


def demo_without_sheet_name():
    """Demonstrate collecting unmatched descriptions without sheet name column"""
    print("=== Demo Without Sheet Name Column ===\n")
    
    # Initialize category dictionary
    category_dict = CategoryDictionary()
    
    # Create test DataFrame without sheet name
    df = create_test_dataframe().drop(columns=['Sheet_Name'])
    print(f"Test DataFrame without sheet name column:")
    print(df[['Description']].head())
    print()
    
    # Perform auto categorization
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.7
    )
    
    # Collect unmatched descriptions without sheet name
    unmatched_descriptions = collect_unmatched_descriptions(
        dataframe=result.dataframe,
        category_column='Category',
        description_column='Description'
        # No sheet_name_column specified
    )
    
    print(f"Unmatched descriptions without sheet name:")
    print(f"  Total unique: {len(unmatched_descriptions)}")
    print()
    
    for i, desc in enumerate(unmatched_descriptions[:3]):
        print(f"{i+1}. '{desc.description}' (frequency: {desc.frequency}, sheet: {desc.source_sheet_name})")
    print()


def demo_export_unmatched():
    """Demonstrate exporting unmatched descriptions"""
    print("=== Export Unmatched Descriptions Demo ===\n")
    
    # Initialize category dictionary
    category_dict = CategoryDictionary()
    
    # Create test DataFrame
    df = create_test_dataframe()
    
    # Perform auto categorization
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.7
    )
    
    # Collect unmatched descriptions
    unmatched_descriptions = collect_unmatched_descriptions(
        dataframe=result.dataframe,
        category_column='Category',
        description_column='Description',
        sheet_name_column='Sheet_Name'
    )
    
    # Export to DataFrame for easy viewing
    export_data = []
    for desc in unmatched_descriptions:
        export_data.append({
            'Description': desc.description,
            'Source_Sheet': desc.source_sheet_name,
            'Frequency': desc.frequency,
            'Sample_Rows': ', '.join(map(str, desc.sample_rows)),
            'Original_Index': desc.original_index
        })
    
    export_df = pd.DataFrame(export_data)
    
    print("Exported unmatched descriptions:")
    print(export_df)
    print()
    
    # Save to CSV
    export_path = Path("examples/unmatched_descriptions.csv")
    export_df.to_csv(export_path, index=False)
    print(f"Exported to: {export_path}")
    print()


def main():
    """Run all demos"""
    # Setup logging
    setup_logging(console_output=True)
    
    print("Collect Unmatched Descriptions Demo Suite")
    print("=" * 60)
    print()
    
    try:
        demo_collect_unmatched_descriptions()
        demo_without_sheet_name()
        demo_export_unmatched()
        
        print("✅ All demos completed successfully!")
        print("\nThe collect_unmatched_descriptions function is ready for integration!")
        
    except Exception as e:
        print(f"❌ Demo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 