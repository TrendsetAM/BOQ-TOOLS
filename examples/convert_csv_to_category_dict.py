#!/usr/bin/env python3
"""
CSV to Category Dictionary Converter
Converts the csvjson.json file to the category dictionary format
"""

import json
import sys
from pathlib import Path
from typing import Dict, List, Set

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary, CategoryMapping


def convert_csv_to_category_dictionary(csv_file_path: Path, output_file_path: Path) -> bool:
    """
    Convert CSV JSON file to category dictionary format
    
    Args:
        csv_file_path: Path to the CSV JSON file
        output_file_path: Path to save the category dictionary
        
    Returns:
        True if conversion was successful
    """
    try:
        # Load CSV JSON data
        with open(csv_file_path, 'r', encoding='utf-8') as f:
            csv_data = json.load(f)
        
        print(f"Loaded {len(csv_data)} items from {csv_file_path}")
        
        # Analyze categories
        categories = set()
        for item in csv_data:
            if 'CATEGORY' in item:
                categories.add(item['CATEGORY'])
        
        print(f"Found {len(categories)} unique categories:")
        for category in sorted(categories):
            print(f"  - {category}")
        print()
        
        # Create category dictionary
        category_dict = CategoryDictionary(dictionary_file=output_file_path)
        
        # Clear existing mappings (since we're creating from scratch)
        category_dict.mappings.clear()
        category_dict.categories.clear()
        
        # Add categories
        for category in categories:
            category_dict.categories.add(category.lower())
        
        # Convert items to mappings
        converted_count = 0
        skipped_count = 0
        
        for item in csv_data:
            if 'CATEGORY' in item and 'DESCRIPTION' in item:
                category = item['CATEGORY'].lower()
                description = item['DESCRIPTION'].strip()
                
                if description:
                    # Create mapping
                    mapping = CategoryMapping(
                        description=description.lower(),
                        category=category,
                        confidence=1.0,
                        notes=f"Converted from CSV data - Category: {item['CATEGORY']}"
                    )
                    
                    # Add to dictionary (avoid duplicates)
                    if description.lower() not in category_dict.mappings:
                        category_dict.mappings[description.lower()] = mapping
                        converted_count += 1
                    else:
                        skipped_count += 1
                        print(f"  Skipped duplicate: '{description[:50]}...'")
            else:
                skipped_count += 1
                print(f"  Skipped item missing required fields: {item}")
        
        print(f"\nConversion Summary:")
        print(f"  Total items processed: {len(csv_data)}")
        print(f"  Successfully converted: {converted_count}")
        print(f"  Skipped: {skipped_count}")
        print(f"  Unique categories: {len(categories)}")
        print(f"  Unique descriptions: {len(category_dict.mappings)}")
        
        # Save the category dictionary
        success = category_dict.save_dictionary()
        if success:
            print(f"\nCategory dictionary saved to: {output_file_path}")
            return True
        else:
            print(f"\nFailed to save category dictionary")
            return False
            
    except Exception as e:
        print(f"Error converting CSV to category dictionary: {e}")
        import traceback
        traceback.print_exc()
        return False


def analyze_csv_data(csv_file_path: Path):
    """Analyze the CSV data structure"""
    try:
        with open(csv_file_path, 'r', encoding='utf-8') as f:
            csv_data = json.load(f)
        
        print(f"CSV Data Analysis for: {csv_file_path}")
        print("=" * 60)
        
        # Basic stats
        print(f"Total items: {len(csv_data)}")
        
        # Check structure
        if csv_data and isinstance(csv_data, list):
            sample_item = csv_data[0]
            print(f"Sample item keys: {list(sample_item.keys())}")
            
            # Analyze categories
            categories = {}
            for item in csv_data:
                if 'CATEGORY' in item:
                    category = item['CATEGORY']
                    categories[category] = categories.get(category, 0) + 1
            
            print(f"\nCategory distribution:")
            for category, count in sorted(categories.items()):
                print(f"  {category}: {count} items")
            
            # Analyze description lengths
            desc_lengths = []
            for item in csv_data:
                if 'DESCRIPTION' in item:
                    desc_lengths.append(len(item['DESCRIPTION']))
            
            if desc_lengths:
                print(f"\nDescription length statistics:")
                print(f"  Average length: {sum(desc_lengths) / len(desc_lengths):.1f} characters")
                print(f"  Min length: {min(desc_lengths)} characters")
                print(f"  Max length: {max(desc_lengths)} characters")
                
                # Show some sample descriptions
                print(f"\nSample descriptions:")
                for i, item in enumerate(csv_data[:5]):
                    if 'DESCRIPTION' in item:
                        desc = item['DESCRIPTION'][:100] + "..." if len(item['DESCRIPTION']) > 100 else item['DESCRIPTION']
                        print(f"  {i+1}. {desc}")
        
    except Exception as e:
        print(f"Error analyzing CSV data: {e}")


def main():
    """Main conversion function"""
    csv_file_path = Path("config/csvjson.json")
    output_file_path = Path("config/category_dictionary.json")
    
    print("CSV to Category Dictionary Converter")
    print("=" * 50)
    print()
    
    # Check if input file exists
    if not csv_file_path.exists():
        print(f"Error: Input file not found: {csv_file_path}")
        return
    
    # Analyze the CSV data first
    analyze_csv_data(csv_file_path)
    print()
    
    # Perform conversion
    print("Starting conversion...")
    success = convert_csv_to_category_dictionary(csv_file_path, output_file_path)
    
    if success:
        print("\n✅ Conversion completed successfully!")
        
        # Test the new category dictionary
        print("\nTesting the new category dictionary...")
        test_category_dictionary(output_file_path)
    else:
        print("\n❌ Conversion failed!")


def test_category_dictionary(dict_file_path: Path):
    """Test the newly created category dictionary"""
    try:
        category_dict = CategoryDictionary(dictionary_file=dict_file_path)
        
        # Test some sample lookups
        test_descriptions = [
            "concrete foundation",
            "electrical installation",
            "site management",
            "crane rental",
            "PV module installation",
            "MV cable supply",
            "tracker installation"
        ]
        
        print("Testing category matching:")
        print("-" * 60)
        
        for desc in test_descriptions:
            match = category_dict.find_category(desc, threshold=0.6)
            status = "✓" if match.matched_category else "✗"
            print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
                  f"(confidence: {match.confidence:.2f}, type: {match.match_type})")
        
        # Show statistics
        stats = category_dict.get_statistics()
        print(f"\nDictionary Statistics:")
        print(f"  Total mappings: {stats['total_mappings']}")
        print(f"  Total categories: {stats['total_categories']}")
        print(f"  Category distribution:")
        for category, count in stats['category_counts'].items():
            print(f"    {category}: {count} mappings")
            
    except Exception as e:
        print(f"Error testing category dictionary: {e}")


if __name__ == "__main__":
    main() 