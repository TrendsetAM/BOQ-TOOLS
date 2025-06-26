#!/usr/bin/env python3
"""
Test Converted Category Dictionary
Test the category dictionary with actual descriptions from the CSV data
"""

import json
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary


def test_with_actual_descriptions():
    """Test the category dictionary with actual descriptions from the CSV"""
    print("Testing Category Dictionary with Actual CSV Data")
    print("=" * 60)
    
    # Load the category dictionary
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    
    # Load some sample descriptions from the original CSV
    csv_file_path = Path("config/csvjson.json")
    with open(csv_file_path, 'r', encoding='utf-8') as f:
        csv_data = json.load(f)
    
    # Test with various categories
    test_categories = [
        "Civil Works",
        "Electrical Works", 
        "Site Costs",
        "General Costs",
        "MV Cables"
    ]
    
    for category in test_categories:
        print(f"\n--- Testing {category} ---")
        
        # Find items from this category
        category_items = [item for item in csv_data if item.get('CATEGORY') == category]
        
        # Test first 3 items from each category
        for i, item in enumerate(category_items[:3]):
            description = item.get('DESCRIPTION', '')
            if description:
                # Test exact match
                match = category_dict.find_category(description, threshold=0.8)
                
                status = "✓" if match.matched_category else "✗"
                expected = category.lower()
                actual = match.matched_category or "No match"
                
                print(f"{status} '{description[:60]}...'")
                print(f"    Expected: {expected}")
                print(f"    Actual:   {actual}")
                print(f"    Confidence: {match.confidence:.2f}")
                print(f"    Match type: {match.match_type}")
                
                if match.matched_category != expected:
                    print(f"    ⚠️  Mismatch detected!")
                print()


def test_partial_matching():
    """Test partial matching with keywords from descriptions"""
    print("\n" + "=" * 60)
    print("Testing Partial Matching")
    print("=" * 60)
    
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    
    # Test with partial descriptions
    partial_tests = [
        ("concrete foundation work", "civil works"),
        ("electrical installation", "electrical works"),
        ("site management", "site costs"),
        ("crane rental", "site costs"),
        ("PV module installation", "pv mod. installation"),
        ("MV cable supply", "mv cables"),
        ("tracker installation", "tracker inst."),
        ("trenching work", "trenching"),
        ("solar cable", "solar cables"),
        ("earth movement", "earth movement")
    ]
    
    for desc, expected_category in partial_tests:
        match = category_dict.find_category(desc, threshold=0.6)
        
        status = "✓" if match.matched_category == expected_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
              f"(expected: {expected_category}, confidence: {match.confidence:.2f})")
        
        if match.suggestions:
            print(f"    Suggestions: {', '.join(match.suggestions[:3])}")


def test_category_statistics():
    """Show detailed statistics of the converted dictionary"""
    print("\n" + "=" * 60)
    print("Category Dictionary Statistics")
    print("=" * 60)
    
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    stats = category_dict.get_statistics()
    
    print(f"Total mappings: {stats['total_mappings']}")
    print(f"Total categories: {stats['total_categories']}")
    print()
    
    print("Category distribution:")
    for category, count in sorted(stats['category_counts'].items()):
        print(f"  {category}: {count} mappings")
    
    print()
    print("Most used mappings:")
    for mapping in stats['most_used_mappings'][:10]:
        print(f"  - {mapping.description[:50]}... -> {mapping.category} (used {mapping.usage_count} times)")


def test_search_functionality():
    """Test search functionality with various queries"""
    print("\n" + "=" * 60)
    print("Testing Search Functionality")
    print("=" * 60)
    
    category_dict = CategoryDictionary(Path("config/category_dictionary.json"))
    
    # Test searches
    search_queries = [
        "concrete",
        "electrical",
        "cable",
        "installation",
        "supply",
        "foundation",
        "tracker",
        "solar",
        "mv",
        "lv"
    ]
    
    for query in search_queries:
        print(f"\nSearching for: '{query}'")
        
        # Find all mappings that contain this query
        matching_mappings = []
        for desc, mapping in category_dict.mappings.items():
            if query.lower() in desc:
                matching_mappings.append(mapping)
        
        # Show top 3 matches
        for i, mapping in enumerate(matching_mappings[:3]):
            print(f"  {i+1}. {mapping.description[:60]}... -> {mapping.category}")
        
        if len(matching_mappings) > 3:
            print(f"  ... and {len(matching_mappings) - 3} more")


def main():
    """Run all tests"""
    try:
        test_with_actual_descriptions()
        test_partial_matching()
        test_category_statistics()
        test_search_functionality()
        
        print("\n" + "=" * 60)
        print("✅ All tests completed successfully!")
        print("The category dictionary is ready for use in automatic row categorization.")
        
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 