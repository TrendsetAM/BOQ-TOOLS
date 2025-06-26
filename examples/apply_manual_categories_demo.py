#!/usr/bin/env python3
"""
Demo script for applying manual categorizations to DataFrame
"""

import sys
import os
from pathlib import Path
import pandas as pd

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.auto_categorizer import CategoryDictionary, auto_categorize_dataset, collect_unmatched_descriptions
from core.manual_categorizer import (
    generate_manual_categorization_excel, 
    process_manual_categorizations,
    apply_manual_categories,
    get_categorization_coverage_report,
    export_categorization_report
)


def main():
    """Demo the apply_manual_categories workflow"""
    print("=== Apply Manual Categories Demo ===\n")
    
    # Load category dictionary
    print("1. Loading category dictionary...")
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    print(f"   Loaded {len(category_dict.mappings)} category mappings")
    print(f"   Available categories: {list(category_dict.get_all_categories())}")
    print()
    
    # Load sample data
    print("2. Loading sample BOQ data...")
    sample_file = "examples/sample_boq.csv"
    if not os.path.exists(sample_file):
        print(f"   Sample file not found: {sample_file}")
        print("   Creating sample data...")
        
        # Create sample data with more variety
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
                'Inverter Setup',  # Duplicate for frequency testing
                'New Unknown Item D',
                'Another Mystery Part E',
                'Test Component F'
            ],
            'Quantity': [10, 5, 100, 20, 1, 15, 8, 12, 5, 3, 25, 7, 9],
            'Unit': ['pcs', 'pcs', 'm', 'sets', 'lot', 'pcs', 'pcs', 'pcs', 'pcs', 'pcs', 'pcs', 'pcs', 'pcs'],
            'Unit_Price': [150.00, 500.00, 2.50, 75.00, 1000.00, 25.00, 45.00, 30.00, 150.00, 500.00, 35.00, 55.00, 40.00]
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
    result = auto_categorize_dataset(df, category_dict)
    categorized_df = result.dataframe
    unmatched_descriptions = result.unmatched_descriptions
    stats = result.match_statistics
    print(f"   Categorized {stats['matched_rows']} out of {stats['total_rows']} descriptions")
    print(f"   Unmatched descriptions: {len(unmatched_descriptions)}")
    print()
    
    # Show initial categorization status
    print("4. Initial categorization status:")
    initial_coverage = get_categorization_coverage_report(categorized_df)
    print(f"   Coverage rate: {initial_coverage['summary']['coverage_rate']:.1%}")
    print(f"   Categorized: {initial_coverage['summary']['categorized_rows']}/{initial_coverage['summary']['total_rows']}")
    print()
    
    # Create sample manual categorizations
    print("5. Creating sample manual categorizations...")
    sample_manual_categorizations = {
        'unknown component a': 'Electrical Components',
        'mystery part b': 'Mechanical Parts',
        'uncategorized item c': 'Tools & Equipment',
        'new unknown item d': 'Materials',
        'another mystery part e': 'Labor',
        'test component f': 'Services'
    }
    print(f"   Created {len(sample_manual_categorizations)} sample categorizations")
    print()
    
    # Apply manual categorizations
    print("6. Applying manual categorizations...")
    try:
        result = apply_manual_categories(
            dataframe=categorized_df,
            manual_categorizations=sample_manual_categorizations,
            description_column='Description',
            category_column='Category',
            case_sensitive=False
        )
        
        updated_df = result['updated_dataframe']
        statistics = result['statistics']
        
        print(f"   ✓ Successfully applied manual categorizations")
        print(f"   Rows updated: {statistics['rows_updated']}")
        print(f"   Exact matches: {statistics['exact_matches']}")
        print(f"   Case-insensitive matches: {statistics['case_insensitive_matches']}")
        print(f"   Coverage rate improved to: {statistics['coverage_rate']:.1%}")
        print()
        
        # Show remaining unmatched
        if result['remaining_unmatched']:
            print(f"   Remaining unmatched descriptions: {len(result['remaining_unmatched'])}")
            print("   Sample remaining unmatched:")
            for desc in result['remaining_unmatched'][:3]:
                print(f"     - '{desc}'")
            print()
        
        # Generate final coverage report
        print("7. Generating final coverage report...")
        final_coverage = get_categorization_coverage_report(updated_df)
        print(f"   Final coverage rate: {final_coverage['summary']['coverage_rate']:.1%}")
        print(f"   Total categorized: {final_coverage['summary']['categorized_rows']}/{final_coverage['summary']['total_rows']}")
        print(f"   Unique categories used: {final_coverage['summary']['unique_categories']}")
        print()
        
        # Show category distribution
        print("   Category distribution:")
        for category, count in final_coverage['top_categories'][:5]:
            print(f"     {category}: {count}")
        print()
        
        # Export coverage report
        print("8. Exporting coverage report...")
        report_file = Path("examples/categorization_coverage_report.xlsx")
        success = export_categorization_report(updated_df, final_coverage, report_file)
        if success:
            print(f"   ✓ Coverage report exported to: {report_file}")
        else:
            print("   ✗ Failed to export coverage report")
        print()
        
        # Save updated data
        print("9. Saving updated data...")
        output_file = "examples/final_categorized_data.csv"
        updated_df.to_csv(output_file, index=False)
        print(f"   ✓ Updated data saved to: {output_file}")
        print()
        
        # Show before/after comparison
        print("10. Before/After Comparison:")
        print(f"   Before: {initial_coverage['summary']['coverage_rate']:.1%} coverage")
        print(f"   After:  {final_coverage['summary']['coverage_rate']:.1%} coverage")
        print(f"   Improvement: {final_coverage['summary']['coverage_rate'] - initial_coverage['summary']['coverage_rate']:.1%}")
        print()
        
    except Exception as e:
        print(f"   Error applying manual categorizations: {e}")
        return
    
    print("=== Demo completed ===")
    print("\nSummary:")
    print(f"- Applied {len(sample_manual_categorizations)} manual categorizations")
    print(f"- Improved coverage from {initial_coverage['summary']['coverage_rate']:.1%} to {final_coverage['summary']['coverage_rate']:.1%}")
    print(f"- Final dataset has {final_coverage['summary']['categorized_rows']} categorized rows")
    print(f"- Used {final_coverage['summary']['unique_categories']} different categories")


def test_case_sensitivity():
    """Test case sensitivity in manual categorization"""
    print("\n=== Testing Case Sensitivity ===\n")
    
    # Create test data
    test_data = {
        'Description': [
            'Solar Panel Installation',
            'SOLAR PANEL INSTALLATION',  # Different case
            'solar panel installation',  # Lower case
            'Solar Panel Installation',  # Same case
            'Unknown Item'
        ],
        'Category': ['', '', '', '', '']
    }
    df = pd.DataFrame(test_data)
    
    # Create manual categorizations with different cases
    manual_cats = {
        'solar panel installation': 'Solar Equipment',  # Lower case
        'unknown item': 'Miscellaneous'
    }
    
    print("Test data:")
    for i, desc in enumerate(test_data['Description']):
        print(f"  {i+1}. '{desc}'")
    print()
    
    print("Manual categorizations:")
    for desc, cat in manual_cats.items():
        print(f"  '{desc}' → '{cat}'")
    print()
    
    # Test case-sensitive matching
    print("Case-sensitive matching:")
    result_sensitive = apply_manual_categories(df, manual_cats, case_sensitive=True)
    print(f"  Rows updated: {result_sensitive['statistics']['rows_updated']}")
    print(f"  Exact matches: {result_sensitive['statistics']['exact_matches']}")
    print()
    
    # Test case-insensitive matching
    print("Case-insensitive matching:")
    result_insensitive = apply_manual_categories(df, manual_cats, case_sensitive=False)
    print(f"  Rows updated: {result_insensitive['statistics']['rows_updated']}")
    print(f"  Exact matches: {result_insensitive['statistics']['exact_matches']}")
    print(f"  Case-insensitive matches: {result_insensitive['statistics']['case_insensitive_matches']}")
    print()
    
    # Show results
    print("Case-sensitive results:")
    for i, row in result_sensitive['updated_dataframe'].iterrows():
        print(f"  '{row['Description']}' → '{row['Category']}'")
    print()
    
    print("Case-insensitive results:")
    for i, row in result_insensitive['updated_dataframe'].iterrows():
        print(f"  '{row['Description']}' → '{row['Category']}'")


if __name__ == "__main__":
    main()
    test_case_sensitivity() 