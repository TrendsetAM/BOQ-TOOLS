#!/usr/bin/env python3
"""
Configuration System Demo
Demonstrates how to use the BOQ configuration system
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from utils.config import get_config, ColumnType


def demo_column_mappings():
    """Demonstrate column mapping functionality"""
    config = get_config()
    
    print("=== Column Mappings Demo ===")
    
    # Show all column types
    print(f"Available column types: {len(config.get_all_column_types())}")
    for col_type in config.get_all_column_types():
        mapping = config.get_column_mapping(col_type)
        if mapping:
            print(f"  {col_type.value}: {len(mapping.keywords)} keywords, "
                  f"required={mapping.required}, weight={mapping.weight}")
    
    # Show required columns
    required = config.get_required_columns()
    print(f"\nRequired columns: {[col.value for col in required]}")
    
    # Demonstrate keyword matching
    test_headers = ["Description", "Qty", "Unit Price", "Total", "Remarks"]
    print(f"\nTesting headers: {test_headers}")
    
    for header in test_headers:
        best_match = None
        best_score = 0
        
        for col_type in config.get_all_column_types():
            mapping = config.get_column_mapping(col_type)
            if mapping:
                for keyword in mapping.keywords:
                    if keyword.lower() in header.lower():
                        score = mapping.weight
                        if score > best_score:
                            best_score = score
                            best_match = col_type.value
        
        if best_match:
            print(f"  '{header}' -> {best_match} (score: {best_score})")
        else:
            print(f"  '{header}' -> no match")


def demo_sheet_classification():
    """Demonstrate sheet classification functionality"""
    config = get_config()
    
    print("\n=== Sheet Classification Demo ===")
    
    # Test different sheet names and content
    test_cases = [
        ("BOQ Main", ["description", "quantity", "unit price", "total"]),
        ("Summary Sheet", ["subtotal", "grand total", "final amount"]),
        ("Preliminaries", ["general items", "site setup", "temporary works"]),
        ("Substructure", ["excavation", "concrete", "foundation"]),
        ("Unknown Sheet", ["random", "data", "here"])
    ]
    
    for sheet_name, content in test_cases:
        sheet_type, confidence = config.get_sheet_classification(sheet_name, content)
        print(f"  '{sheet_name}' -> {sheet_type} (confidence: {confidence:.2f})")


def demo_validation_thresholds():
    """Demonstrate validation thresholds"""
    config = get_config()
    
    print("\n=== Validation Thresholds Demo ===")
    
    thresholds = config.validation_thresholds
    print(f"Min column confidence: {thresholds.min_column_confidence}")
    print(f"Min sheet confidence: {thresholds.min_sheet_confidence}")
    print(f"Max empty rows %: {thresholds.max_empty_rows_percentage}")
    print(f"Min data rows: {thresholds.min_data_rows}")
    print(f"Max header rows: {thresholds.max_header_rows}")


def demo_processing_limits():
    """Demonstrate processing limits"""
    config = get_config()
    
    print("\n=== Processing Limits Demo ===")
    
    limits = config.processing_limits
    print(f"Max file size: {limits.max_file_size_mb} MB")
    print(f"Max sheets per file: {limits.max_sheets_per_file}")
    print(f"Max rows per sheet: {limits.max_rows_per_sheet}")
    print(f"Max columns per sheet: {limits.max_columns_per_sheet}")
    print(f"Timeout: {limits.timeout_seconds} seconds")
    print(f"Memory limit: {limits.memory_limit_mb} MB")


def demo_export_settings():
    """Demonstrate export settings"""
    config = get_config()
    
    print("\n=== Export Settings Demo ===")
    
    export = config.export_settings
    print(f"Output format: {export.output_format}")
    print(f"Include summary: {export.include_summary}")
    print(f"Include validation report: {export.include_validation_report}")
    print(f"Sheet name template: {export.sheet_name_template}")
    print(f"Backup original: {export.backup_original}")
    print(f"Compression level: {export.compression_level}")


def main():
    """Run all demos"""
    print("BOQ Configuration System Demo")
    print("=" * 40)
    
    demo_column_mappings()
    demo_sheet_classification()
    demo_validation_thresholds()
    demo_processing_limits()
    demo_export_settings()
    
    print("\n" + "=" * 40)
    print("Demo completed successfully!")


if __name__ == "__main__":
    main() 