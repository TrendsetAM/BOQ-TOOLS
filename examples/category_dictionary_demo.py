#!/usr/bin/env python3
"""
Category Dictionary Demo
Demonstrates the CategoryDictionary functionality for automatic row categorization
"""

import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary, CategoryType, CategoryMapping
from utils.logger import setup_logging


def demo_basic_functionality():
    """Demonstrate basic CategoryDictionary functionality"""
    print("=== Category Dictionary Basic Functionality Demo ===\n")
    
    # Initialize category dictionary
    category_dict = CategoryDictionary()
    
    # Show available categories
    print("Available categories:")
    categories = category_dict.get_all_categories()
    for i, category in enumerate(categories, 1):
        print(f"  {i}. {category}")
    print()
    
    # Test category matching
    test_descriptions = [
        "concrete foundation",
        "steel reinforcement",
        "electrical installation",
        "crane rental",
        "site office setup",
        "unknown material",
        "excavation work",
        "painting services"
    ]
    
    print("Testing category matching:")
    print("-" * 60)
    
    for desc in test_descriptions:
        match = category_dict.find_category(desc)
        status = "✓" if match.matched_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
              f"(confidence: {match.confidence:.2f}, type: {match.match_type})")
        
        if match.suggestions:
            print(f"    Suggestions: {', '.join(match.suggestions[:3])}")
        print()
    
    # Show statistics
    print("Dictionary Statistics:")
    stats = category_dict.get_statistics()
    print(f"  Total mappings: {stats['total_mappings']}")
    print(f"  Total categories: {stats['total_categories']}")
    print(f"  Category distribution:")
    for category, count in stats['category_counts'].items():
        print(f"    {category}: {count} mappings")
    print()


def demo_adding_mappings():
    """Demonstrate adding new mappings"""
    print("=== Adding New Mappings Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Add new mappings
    new_mappings = [
        ("aluminum cladding", "materials"),
        ("HVAC installation", "labor"),
        ("scaffolding rental", "equipment"),
        ("soil testing", "services"),
        ("project management", "overhead")
    ]
    
    print("Adding new mappings:")
    for desc, category in new_mappings:
        success = category_dict.add_mapping(desc, category)
        status = "✓" if success else "✗"
        print(f"{status} Added: '{desc}' -> '{category}'")
    
    print()
    
    # Test the new mappings
    print("Testing new mappings:")
    for desc, expected_category in new_mappings:
        match = category_dict.find_category(desc)
        status = "✓" if match.matched_category == expected_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category} "
              f"(expected: {expected_category})")
    
    print()
    
    # Save the updated dictionary
    success = category_dict.save_dictionary()
    print(f"Dictionary saved: {'✓' if success else '✗'}")
    print()


def demo_partial_matching():
    """Demonstrate partial matching functionality"""
    print("=== Partial Matching Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Test partial matches
    partial_tests = [
        ("concrete foundation work", "concrete"),
        ("steel reinforcement bars", "steel"),
        ("electrical wiring installation", "electrical"),
        ("crane lifting services", "crane"),
        ("site office construction", "site office")
    ]
    
    print("Testing partial matching:")
    print("-" * 50)
    
    for desc, expected_keyword in partial_tests:
        match = category_dict.find_category(desc, threshold=0.6)
        status = "✓" if match.matched_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
              f"(confidence: {match.confidence:.2f}, type: {match.match_type})")
        
        if match.match_type == 'partial':
            print(f"    Partial match with keyword: '{expected_keyword}'")
        print()
    
    print()


def demo_fuzzy_matching():
    """Demonstrate fuzzy matching functionality"""
    print("=== Fuzzy Matching Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Test fuzzy matches (typos, variations)
    fuzzy_tests = [
        ("concret", "concrete"),  # Missing 'e'
        ("steel reinforcment", "steel"),  # Missing 'e'
        ("electrical instalation", "electrical"),  # Missing 'l'
        ("crane rentel", "crane"),  # Typo
        ("site offce", "site office")  # Missing 'i'
    ]
    
    print("Testing fuzzy matching:")
    print("-" * 50)
    
    for desc, expected_keyword in fuzzy_tests:
        match = category_dict.find_category(desc, threshold=0.7)
        status = "✓" if match.matched_category else "✗"
        print(f"{status} '{desc}' -> {match.matched_category or 'No match'} "
              f"(confidence: {match.confidence:.2f}, type: {match.match_type})")
        
        if match.match_type == 'fuzzy':
            print(f"    Fuzzy match with keyword: '{expected_keyword}'")
        print()
    
    print()


def demo_category_management():
    """Demonstrate category management functions"""
    print("=== Category Management Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Get mappings for a specific category
    print("Mappings for 'materials' category:")
    material_mappings = category_dict.get_mappings_for_category("materials")
    for mapping in material_mappings[:5]:  # Show first 5
        print(f"  - {mapping.description} (usage: {mapping.usage_count})")
    print()
    
    # Update a mapping
    print("Updating mapping:")
    success = category_dict.update_mapping("concrete", "materials", new_notes="Updated for demo")
    print(f"  Update 'concrete' mapping: {'✓' if success else '✗'}")
    
    # Remove a mapping
    print("Removing mapping:")
    success = category_dict.remove_mapping("test_mapping")
    print(f"  Remove 'test_mapping': {'✓' if success else '✗'}")
    print()
    
    # Show most used mappings
    print("Most used mappings:")
    stats = category_dict.get_statistics()
    for mapping in stats['most_used_mappings'][:5]:
        print(f"  - {mapping.description} -> {mapping.category} (used {mapping.usage_count} times)")
    print()


def demo_export_import():
    """Demonstrate export and import functionality"""
    print("=== Export/Import Demo ===\n")
    
    category_dict = CategoryDictionary()
    
    # Export dictionary
    export_path = Path("examples/exported_category_dict.json")
    success = category_dict.export_dictionary(export_path)
    print(f"Export dictionary: {'✓' if success else '✗'}")
    
    # Create new dictionary and import
    new_dict = CategoryDictionary()
    success = new_dict.import_dictionary(export_path)
    print(f"Import dictionary: {'✓' if success else '✗'}")
    
    # Compare statistics
    original_stats = category_dict.get_statistics()
    imported_stats = new_dict.get_statistics()
    
    print(f"Original mappings: {original_stats['total_mappings']}")
    print(f"Imported mappings: {imported_stats['total_mappings']}")
    print(f"Import successful: {'✓' if original_stats['total_mappings'] == imported_stats['total_mappings'] else '✗'}")
    
    # Clean up
    if export_path.exists():
        export_path.unlink()
    print()


def main():
    """Run all demos"""
    # Setup logging
    setup_logging(console_output=True)
    
    print("Category Dictionary Demo Suite")
    print("=" * 50)
    print()
    
    try:
        demo_basic_functionality()
        demo_adding_mappings()
        demo_partial_matching()
        demo_fuzzy_matching()
        demo_category_management()
        demo_export_import()
        
        print("All demos completed successfully!")
        
    except Exception as e:
        print(f"Demo failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main() 