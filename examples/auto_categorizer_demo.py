#!/usr/bin/env python3
"""
Auto Categorizer Demo
Demonstrates the auto_categorize_dataset function with sample data
"""

import sys
import pandas as pd
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary
from core.auto_categorizer import auto_categorize_dataset
from utils.logger import setup_logging


def create_sample_dataframe():
    """Create a sample DataFrame for testing"""
    sample_data = {
        'Description': [
            'Concrete foundation work',
            'Steel reinforcement supply',
            'Electrical installation',
            'Crane rental for lifting',
            'Site office setup',
            'Project management',
            'MV cable installation',
            'Transformer installation',
            'Unknown material type',
            'Custom electrical work',
            'Site camp construction',
            'Utilities and power supply',
            'Permits and documentation',
            'Insurance and bonds',
            'Road construction',
            'Foundation excavation',
            'Testing and certification',
            'Design and engineering',
            'Plumbing installation',
            'Carpentry work'
        ],
        'Quantity': [100, 50, 200, 10, 1, 1, 500, 5, 25, 150, 1, 1, 1, 1, 200, 300, 1, 1, 100, 80],
        'Unit': ['m3', 'tons', 'm', 'days', 'unit', 'month', 'm', 'units', 'units', 'm', 'unit', 'unit', 'unit', 'unit', 'm2', 'm3', 'unit', 'unit', 'm', 'm2'],
        'Unit_Price': [150, 800, 25, 500, 5000, 8000, 45, 15000, 100, 30, 20000, 15000, 5000, 10000, 80, 60, 3000, 12000, 35, 45]
    }
    
    return pd.DataFrame(sample_data)


def demo_basic_categorization():
    """Demonstrate basic auto categorization"""
    print("=== Auto Categorization Demo ===\n")
    
    # Initialize category dictionary
    category_dict = CategoryDictionary()
    
    # Create sample DataFrame
    df = create_sample_dataframe()
    print(f"Sample DataFrame with {len(df)} rows:")
    print(df.head())
    print()
    
    # Define progress callback
    def progress_callback(percentage, message):
        print(f"  {percentage:.1f}% - {message}")
    
    # Perform auto categorization
    print("Starting auto categorization...")
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.7,
        progress_callback=progress_callback
    )
    
    print("\nCategorization Results:")
    print(f"  Total rows: {result.total_rows}")
    print(f"  Matched rows: {result.matched_rows} ({result.match_rate:.1%})")
    print(f"  Unmatched rows: {result.unmatched_rows}")
    print()
    
    # Show categorized DataFrame
    print("Categorized DataFrame:")
    print(result.dataframe[['Description', 'Category']].head(10))
    print()
    
    # Show category distribution
    print("Category Distribution:")
    category_dist = result.match_statistics['category_distribution']
    for category, count in sorted(category_dist.items()):
        if category:  # Skip empty categories
            print(f"  {category}: {count} items")
    print()
    
    # Show unmatched descriptions
    if result.unmatched_descriptions:
        print(f"Unmatched Descriptions ({len(result.unmatched_descriptions)}):")
        for desc in result.unmatched_descriptions[:5]:
            print(f"  - {desc}")
        if len(result.unmatched_descriptions) > 5:
            print(f"  ... and {len(result.unmatched_descriptions) - 5} more")
    print()


def demo_match_types():
    """Demonstrate different match types"""
    print("=== Match Types Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Test different types of descriptions
    test_descriptions = [
        "concrete foundation work",  # Should be exact match
        "concrete foundation",       # Should be partial match
        "concret foundation",        # Should be fuzzy match (typo)
        "electrical installation",   # Should be exact match
        "electrical work",           # Should be partial match
        "unknown material",          # Should be no match
        "steel reinforcement",       # Should be exact match
        "steel reinforcment",        # Should be fuzzy match (typo)
        "crane rental",              # Should be exact match
        "crane rentel"               # Should be fuzzy match (typo)
    ]
    
    print("Testing different match types:")
    print("-" * 60)
    
    for desc in test_descriptions:
        match = category_dict.find_category(desc, threshold=0.6)
        status = "✓" if match.matched_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
              f"(confidence: {match.confidence:.2f}, type: {match.match_type})")
    
    print()


def demo_confidence_thresholds():
    """Demonstrate different confidence thresholds"""
    print("=== Confidence Thresholds Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Test description with different thresholds
    test_description = "concrete foundation work"
    
    print(f"Testing description: '{test_description}'")
    print("-" * 50)
    
    thresholds = [0.9, 0.8, 0.7, 0.6, 0.5]
    
    for threshold in thresholds:
        match = category_dict.find_category(test_description, threshold)
        status = "✓" if match.matched_category else "✗"
        print(f"Threshold {threshold}: {status} -> {match.matched_category or 'No match'} "
              f"(confidence: {match.confidence:.2f})")
    
    print()


def demo_with_real_data():
    """Demonstrate with data from the category dictionary"""
    print("=== Real Data Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Get some real descriptions from the dictionary
    real_descriptions = []
    for desc, mapping in list(category_dict.mappings.items())[:10]:
        real_descriptions.append(mapping.description)
    
    # Create DataFrame with real descriptions
    df = pd.DataFrame({
        'Description': real_descriptions,
        'Quantity': [100] * len(real_descriptions),
        'Unit': ['units'] * len(real_descriptions)
    })
    
    print(f"Testing with {len(df)} real descriptions from the dictionary")
    
    # Perform categorization
    result = auto_categorize_dataset(
        dataframe=df,
        category_dictionary=category_dict,
        description_column='Description',
        category_column='Category',
        confidence_threshold=0.8
    )
    
    print(f"\nResults:")
    print(f"  Match rate: {result.match_rate:.1%}")
    print(f"  Categories found: {result.match_statistics['unique_categories_found']}")
    
    # Show some results
    print(f"\nSample categorized data:")
    print(result.dataframe[['Description', 'Category']].head())
    print()


def main():
    """Run all demos"""
    # Setup logging
    setup_logging(console_output=True)
    
    print("Auto Categorizer Demo Suite")
    print("=" * 50)
    print()
    
    try:
        demo_basic_categorization()
        demo_match_types()
        demo_confidence_thresholds()
        demo_with_real_data()
        
        print("✅ All demos completed successfully!")
        print("\nThe auto_categorize_dataset function is ready for integration!")
        
    except Exception as e:
        print(f"❌ Demo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 