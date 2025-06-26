#!/usr/bin/env python3
"""
Excel to Category Dictionary Converter
Converts Excel files with Category and Description columns to the category dictionary format
"""

import json
import sys
import pandas as pd
from pathlib import Path
from typing import Dict, List, Set, Optional

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.category_dictionary import CategoryDictionary, CategoryMapping


def convert_excel_to_category_dictionary(excel_file_path: Path, output_file_path: Path, 
                                       category_column: str = "Category", 
                                       description_column: str = "Description") -> bool:
    """
    Convert Excel file to category dictionary format
    
    Args:
        excel_file_path: Path to the Excel file
        output_file_path: Path to save the category dictionary
        category_column: Name of the category column
        description_column: Name of the description column
        
    Returns:
        True if conversion was successful
    """
    try:
        # Load Excel file
        print(f"Loading Excel file: {excel_file_path}")
        
        # Try to read the Excel file
        try:
            df = pd.read_excel(excel_file_path)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            print("Trying with different engine...")
            try:
                df = pd.read_excel(excel_file_path, engine='openpyxl')
            except Exception as e2:
                print(f"Error with openpyxl engine: {e2}")
                return False
        
        print(f"Successfully loaded Excel file with {len(df)} rows and {len(df.columns)} columns")
        print(f"Columns found: {list(df.columns)}")
        
        # Check if required columns exist
        if category_column not in df.columns:
            print(f"Error: Category column '{category_column}' not found in Excel file")
            print(f"Available columns: {list(df.columns)}")
            return False
            
        if description_column not in df.columns:
            print(f"Error: Description column '{description_column}' not found in Excel file")
            print(f"Available columns: {list(df.columns)}")
            return False
        
        # Show sample data
        print(f"\nSample data (first 5 rows):")
        print(df[[category_column, description_column]].head())
        print()
        
        # Analyze categories
        categories = df[category_column].dropna().unique()
        print(f"Found {len(categories)} unique categories:")
        for category in sorted(categories):
            count = len(df[df[category_column] == category])
            print(f"  - {category}: {count} items")
        print()
        
        # Create category dictionary
        category_dict = CategoryDictionary(dictionary_file=output_file_path)
        
        # Clear existing mappings (since we're creating from scratch)
        category_dict.mappings.clear()
        category_dict.categories.clear()
        
        # Add categories
        for category in categories:
            if pd.notna(category) and str(category).strip():
                category_dict.categories.add(str(category).lower().strip())
        
        # Convert items to mappings
        converted_count = 0
        skipped_count = 0
        duplicate_count = 0
        
        for index, row in df.iterrows():
            category = row[category_column]
            description = row[description_column]
            
            # Skip rows with missing data
            if pd.isna(category) or pd.isna(description):
                skipped_count += 1
                continue
                
            category_str = str(category).strip()
            description_str = str(description).strip()
            
            if not category_str or not description_str:
                skipped_count += 1
                continue
            
            # Create mapping
            mapping = CategoryMapping(
                description=description_str.lower(),
                category=category_str.lower(),
                confidence=1.0,
                notes=f"Converted from Excel data - Category: {category_str}"
            )
            
            # Add to dictionary (avoid duplicates)
            if description_str.lower() not in category_dict.mappings:
                category_dict.mappings[description_str.lower()] = mapping
                converted_count += 1
            else:
                duplicate_count += 1
                print(f"  Skipped duplicate: '{description_str[:50]}...'")
        
        print(f"\nConversion Summary:")
        print(f"  Total rows in Excel: {len(df)}")
        print(f"  Successfully converted: {converted_count}")
        print(f"  Skipped (missing data): {skipped_count}")
        print(f"  Duplicates skipped: {duplicate_count}")
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
        print(f"Error converting Excel to category dictionary: {e}")
        import traceback
        traceback.print_exc()
        return False


def analyze_excel_file(excel_file_path: Path):
    """Analyze the Excel file structure"""
    try:
        print(f"Excel File Analysis for: {excel_file_path}")
        print("=" * 60)
        
        # Load Excel file
        try:
            df = pd.read_excel(excel_file_path)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            print("Trying with openpyxl engine...")
            df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        # Basic stats
        print(f"Total rows: {len(df)}")
        print(f"Total columns: {len(df.columns)}")
        print(f"Column names: {list(df.columns)}")
        
        # Check for Category and Description columns
        category_cols = [col for col in df.columns if 'category' in col.lower()]
        description_cols = [col for col in df.columns if 'description' in col.lower()]
        
        print(f"\nPotential category columns: {category_cols}")
        print(f"Potential description columns: {description_cols}")
        
        # Show data types
        print(f"\nData types:")
        for col in df.columns:
            print(f"  {col}: {df[col].dtype}")
        
        # Show sample data
        print(f"\nSample data (first 3 rows):")
        print(df.head(3))
        
        # Check for missing values
        print(f"\nMissing values per column:")
        for col in df.columns:
            missing_count = df[col].isna().sum()
            if missing_count > 0:
                print(f"  {col}: {missing_count} missing values")
        
        # If we have category and description columns, analyze them
        if category_cols and description_cols:
            category_col = category_cols[0]
            description_col = description_cols[0]
            
            print(f"\nAnalyzing {category_col} and {description_col}:")
            
            # Category analysis
            categories = df[category_col].dropna().unique()
            print(f"  Unique categories: {len(categories)}")
            for category in sorted(categories)[:10]:  # Show first 10
                count = len(df[df[category_col] == category])
                print(f"    {category}: {count} items")
            
            if len(categories) > 10:
                print(f"    ... and {len(categories) - 10} more categories")
            
            # Description analysis
            desc_lengths = df[description_col].dropna().str.len()
            print(f"\n  Description length statistics:")
            print(f"    Average length: {desc_lengths.mean():.1f} characters")
            print(f"    Min length: {desc_lengths.min()} characters")
            print(f"    Max length: {desc_lengths.max()} characters")
            
            # Show sample descriptions
            print(f"\n  Sample descriptions:")
            for i, desc in enumerate(df[description_col].dropna().head(3)):
                short_desc = desc[:100] + "..." if len(desc) > 100 else desc
                print(f"    {i+1}. {short_desc}")
        
    except Exception as e:
        print(f"Error analyzing Excel file: {e}")


def main():
    """Main conversion function"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert Excel file to Category Dictionary")
    parser.add_argument("excel_file", help="Path to Excel file")
    parser.add_argument("--output", "-o", default="config/category_dictionary.json", 
                       help="Output file path (default: config/category_dictionary.json)")
    parser.add_argument("--category-col", default="Category", 
                       help="Name of category column (default: Category)")
    parser.add_argument("--description-col", default="Description", 
                       help="Name of description column (default: Description)")
    parser.add_argument("--analyze-only", action="store_true", 
                       help="Only analyze the Excel file, don't convert")
    
    args = parser.parse_args()
    
    excel_file_path = Path(args.excel_file)
    output_file_path = Path(args.output)
    
    print("Excel to Category Dictionary Converter")
    print("=" * 50)
    print()
    
    # Check if input file exists
    if not excel_file_path.exists():
        print(f"Error: Input file not found: {excel_file_path}")
        return
    
    # Analyze the Excel file first
    analyze_excel_file(excel_file_path)
    print()
    
    if args.analyze_only:
        print("Analysis complete. Use --help to see conversion options.")
        return
    
    # Perform conversion
    print("Starting conversion...")
    success = convert_excel_to_category_dictionary(
        excel_file_path, 
        output_file_path,
        args.category_col,
        args.description_col
    )
    
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
        
        # Show statistics
        stats = category_dict.get_statistics()
        print(f"\nDictionary Statistics:")
        print(f"  Total mappings: {stats['total_mappings']}")
        print(f"  Total categories: {stats['total_categories']}")
        print(f"  Category distribution:")
        for category, count in sorted(stats['category_counts'].items()):
            print(f"    {category}: {count} mappings")
        
        # Test some sample lookups
        print(f"\nTesting sample lookups:")
        sample_mappings = list(category_dict.mappings.values())[:5]
        for mapping in sample_mappings:
            match = category_dict.find_category(mapping.description, threshold=0.8)
            status = "✓" if match.matched_category else "✗"
            print(f"{status} '{mapping.description[:50]}...' -> {match.matched_category}")
            
    except Exception as e:
        print(f"Error testing category dictionary: {e}")


if __name__ == "__main__":
    main() 