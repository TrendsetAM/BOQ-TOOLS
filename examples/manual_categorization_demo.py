#!/usr/bin/env python3
"""
Demo script for manual categorization Excel generation
"""

import sys
import os
from pathlib import Path
import pandas as pd

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.auto_categorizer import CategoryDictionary, auto_categorize_dataset, collect_unmatched_descriptions
from core.manual_categorizer import generate_manual_categorization_excel, load_manual_categorization_results, apply_manual_categorizations, create_categorization_summary


def main():
    """Demo the manual categorization workflow"""
    print("=== Manual Categorization Demo ===\n")
    
    # Load category dictionary
    print("1. Loading category dictionary...")
    category_dict = CategoryDictionary()
    category_dict.load_from_json("config/category_dictionary.json")
    print(f"   Loaded {len(category_dict)} category mappings")
    print(f"   Available categories: {list(category_dict.get_all_categories())}")
    print()
    
    # Load sample data
    print("2. Loading sample BOQ data...")
    sample_file = "examples/sample_boq.csv"
    if not os.path.exists(sample_file):
        print(f"   Sample file not found: {sample_file}")
        print("   Creating sample data...")
        
        # Create sample data
        sample_data = {
            'Description': [
                'Solar Panel Installation',
                'Inverter Setup',
                'Cable Management',
                'Mounting System',
                'Electrical Testing',
                'Unknown Component A',
                'Mystery Part B',
                'Uncategorized Item C',
                'Solar Panel Installation',  # Duplicate for frequency testing
                'Inverter Setup'  # Duplicate for frequency testing
            ],
            'Quantity': [10, 5, 100, 20, 1, 15, 8, 12, 5, 3],
            'Unit': ['pcs', 'pcs', 'm', 'sets', 'lot', 'pcs', 'pcs', 'pcs', 'pcs', 'pcs'],
            'Unit_Price': [150.00, 500.00, 2.50, 75.00, 1000.00, 25.00, 45.00, 30.00, 150.00, 500.00]
        }
        df = pd.DataFrame(sample_data)
        df.to_csv(sample_file, index=False)
        print(f"   Created sample file: {sample_file}")
    else:
        df = pd.read_csv(sample_file)
        print(f"   Loaded {len(df)} rows from {sample_file}")
    print()
    
    # Auto-categorize the data
    print("3. Auto-categorizing data...")
    categorized_df, unmatched_descriptions, stats = auto_categorize_dataset(df, category_dict)
    print(f"   Categorized {stats['categorized_count']} out of {stats['total_count']} descriptions")
    print(f"   Unmatched descriptions: {len(unmatched_descriptions)}")
    print()
    
    # Collect unmatched descriptions
    print("4. Collecting unmatched descriptions...")
    unmatched_list = collect_unmatched_descriptions(categorized_df, 'Description')
    print(f"   Collected {len(unmatched_list)} unique unmatched descriptions")
    print()
    
    # Generate manual categorization Excel file
    print("5. Generating manual categorization Excel file...")
    available_categories = list(category_dict.get_all_categories())
    
    try:
        excel_filepath = generate_manual_categorization_excel(
            unmatched_descriptions=unmatched_list,
            available_categories=available_categories,
            output_dir=Path("examples")
        )
        print(f"   Excel file created: {excel_filepath}")
        print("   Please open this file and manually categorize the descriptions.")
        print("   Then save the file and run this demo again to load the results.")
        print()
        
        # Check if the file has been manually edited (has categories filled in)
        if excel_filepath.exists():
            try:
                manual_results = load_manual_categorization_results(excel_filepath)
                if manual_results:
                    print("6. Loading manual categorization results...")
                    print(f"   Found {len(manual_results)} manual categorizations")
                    
                    # Apply manual categorizations
                    print("7. Applying manual categorizations...")
                    updated_df = apply_manual_categorizations(
                        categorized_df, 
                        manual_results, 
                        'Description', 
                        'Category'
                    )
                    
                    # Create summary
                    print("8. Creating categorization summary...")
                    summary = create_categorization_summary(updated_df, 'Category')
                    print(f"   Total rows: {summary['total_rows']}")
                    print(f"   Categorized rows: {summary['categorized_rows']}")
                    print(f"   Uncategorized rows: {summary['uncategorized_rows']}")
                    print(f"   Categorization rate: {summary['categorization_rate']:.1%}")
                    print(f"   Unique categories used: {summary['unique_categories']}")
                    print()
                    
                    # Show category distribution
                    print("   Category distribution:")
                    for category, count in summary['category_distribution'].items():
                        print(f"     {category}: {count}")
                    print()
                    
                    # Save updated data
                    output_file = "examples/updated_categorized_data.csv"
                    updated_df.to_csv(output_file, index=False)
                    print(f"9. Updated data saved to: {output_file}")
                    
                else:
                    print("6. No manual categorizations found in the Excel file.")
                    print("   Please open the Excel file, categorize the descriptions,")
                    print("   save the file, and run this demo again.")
                    
            except Exception as e:
                print(f"6. Error loading manual results: {e}")
                print("   This is expected if the file hasn't been manually edited yet.")
        
    except Exception as e:
        print(f"   Error generating Excel file: {e}")
        return
    
    print("\n=== Demo completed ===")
    print("\nNext steps:")
    print("1. Open the generated Excel file")
    print("2. Go to the 'Categorization' sheet")
    print("3. Select categories from the dropdown for each description")
    print("4. Add any notes if needed")
    print("5. Save the file")
    print("6. Run this demo again to load and apply the manual categorizations")


if __name__ == "__main__":
    main() 