#!/usr/bin/env python3
"""
Column Mapper Demo
Demonstrates intelligent column mapping and header detection
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.column_mapper import ColumnMapper, map_columns_quick
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')


def create_sample_sheets():
    """Create sample sheet data for testing"""
    sheets = {
        "Standard BOQ": [
            ["Item Code", "Description", "Unit", "Quantity", "Unit Price", "Total Amount"],
            ["001", "Excavation for foundation", "m³", "100", "25.50", "2550.00"],
            ["002", "Concrete foundation", "m³", "50", "150.00", "7500.00"],
            ["003", "Reinforcement steel", "kg", "2000", "2.50", "5000.00"]
        ],
        
        "Merged Headers": [
            ["", "", "Quantity", "", ""],
            ["Code", "Work Description", "Unit", "Rate", "Amount"],
            ["001", "Excavation", "m³", "25.50", "2550.00"],
            ["002", "Concrete", "m³", "150.00", "7500.00"]
        ],
        
        "Multi-row Headers": [
            ["Item", "Work Description", "Measurement", "Pricing"],
            ["Code", "Scope of Work", "Unit", "Rate", "Total"],
            ["001", "Excavation works", "m³", "25.50", "2550.00"],
            ["002", "Concrete works", "m³", "150.00", "7500.00"]
        ],
        
        "Ambiguous Headers": [
            ["Ref", "Details", "No", "Cost", "Value"],
            ["001", "Excavation", "100", "25.50", "2550.00"],
            ["002", "Concrete", "50", "150.00", "7500.00"]
        ],
        
        "No Headers": [
            ["001", "Excavation", "m³", "100", "25.50", "2550.00"],
            ["002", "Concrete", "m³", "50", "150.00", "7500.00"],
            ["003", "Steel", "kg", "2000", "2.50", "5000.00"]
        ],
        
        "Mixed Content": [
            ["Project Information", "", "", "", ""],
            ["Item Code", "Description", "Unit", "Qty", "Rate", "Amount"],
            ["001", "Excavation", "m³", "100", "25.50", "2550.00"],
            ["002", "Concrete", "m³", "50", "150.00", "7500.00"],
            ["", "Subtotal", "", "", "", "10050.00"],
            ["", "Contingency (10%)", "", "", "", "1005.00"],
            ["", "Total", "", "", "", "11055.00"]
        ]
    }
    
    return sheets


def demo_header_detection():
    """Demonstrate header row detection"""
    print("=== Header Row Detection ===")
    
    mapper = ColumnMapper()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nDetecting headers: {sheet_name}")
        print("-" * 40)
        
        header_info = mapper.find_header_row(sheet_data)
        
        print(f"Header row index: {header_info.row_index}")
        print(f"Confidence: {header_info.confidence:.2f}")
        print(f"Method: {header_info.method.value}")
        print(f"Is merged: {header_info.is_merged}")
        
        print("Reasoning:")
        for reason in header_info.reasoning:
            print(f"  - {reason}")
        
        print("Headers:")
        for i, header in enumerate(header_info.headers):
            print(f"  Column {i + 1}: '{header}'")


def demo_column_mapping():
    """Demonstrate column mapping"""
    print("\n=== Column Mapping ===")
    
    mapper = ColumnMapper()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nMapping columns: {sheet_name}")
        print("-" * 40)
        
        # Find header row first
        header_info = mapper.find_header_row(sheet_data)
        
        if header_info and header_info.headers:
            # Map columns
            mappings = mapper.map_columns_to_types(header_info.headers)
            
            print(f"Found {len(mappings)} column mappings:")
            for mapping in mappings:
                print(f"  Column {mapping.column_index + 1}: '{mapping.original_header}' -> {mapping.mapped_type.value}")
                print(f"    Confidence: {mapping.confidence:.2f}")
                print(f"    Normalized: '{mapping.normalized_header}'")
                
                if mapping.alternatives:
                    print(f"    Alternatives: {[(alt[0].value, alt[1]) for alt in mapping.alternatives[:2]]}")
                
                print(f"    Reasoning: {mapping.reasoning[0] if mapping.reasoning else 'None'}")
        else:
            print("  No headers found")


def demo_complete_mapping():
    """Demonstrate complete mapping process"""
    print("\n=== Complete Mapping Process ===")
    
    mapper = ColumnMapper()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nComplete mapping: {sheet_name}")
        print("-" * 40)
        
        result = mapper.process_sheet_mapping(sheet_data)
        
        print(f"Header row: {result.header_row.row_index} (confidence: {result.header_row.confidence:.2f})")
        print(f"Overall confidence: {result.overall_confidence:.2f}")
        print(f"Mapped columns: {len(result.mappings)}")
        print(f"Unmapped columns: {len(result.unmapped_columns)}")
        
        if result.mappings:
            print("Mappings:")
            for mapping in result.mappings:
                print(f"  {mapping.column_index + 1}: {mapping.mapped_type.value} ({mapping.confidence:.2f})")
        
        if result.unmapped_columns:
            print("Unmapped columns:")
            for col_idx in result.unmapped_columns:
                if col_idx < len(result.header_row.headers):
                    header = result.header_row.headers[col_idx]
                    print(f"  Column {col_idx + 1}: '{header}'")
        
        if result.suggestions:
            print("Suggestions:")
            for suggestion in result.suggestions:
                print(f"  - {suggestion}")


def demo_alternative_mappings():
    """Demonstrate alternative mappings for ambiguous cases"""
    print("\n=== Alternative Mappings ===")
    
    mapper = ColumnMapper()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nAlternatives: {sheet_name}")
        print("-" * 30)
        
        header_info = mapper.find_header_row(sheet_data)
        if not header_info or not header_info.headers:
            continue
        
        alternatives = mapper.get_alternative_mappings(header_info.headers)
        
        if alternatives:
            for col_idx, alt_list in alternatives.items():
                if col_idx < len(header_info.headers):
                    header = header_info.headers[col_idx]
                    print(f"Column {col_idx + 1} '{header}' alternatives:")
                    for col_type, score in alt_list:
                        print(f"  - {col_type.value}: {score:.2f}")
        else:
            print("  No alternative mappings found")


def demo_quick_mapping():
    """Demonstrate quick mapping function"""
    print("\n=== Quick Mapping ===")
    
    # Test with various header sets
    test_headers = [
        ["Description", "Qty", "Unit Price", "Total"],
        ["Item Code", "Work Description", "Unit", "Quantity", "Rate", "Amount"],
        ["Ref", "Details", "No", "Cost", "Value"],
        ["Code", "Description", "Unit", "Qty", "Rate", "Total"]
    ]
    
    for i, headers in enumerate(test_headers, 1):
        print(f"\nTest {i}: {headers}")
        print("-" * 30)
        
        mappings = map_columns_quick(headers)
        
        for col_idx, col_type in mappings.items():
            print(f"  Column {col_idx + 1}: {col_type}")


def demo_normalization():
    """Demonstrate header text normalization"""
    print("\n=== Header Normalization ===")
    
    mapper = ColumnMapper()
    
    test_headers = [
        "Item Code",
        "Work Description",
        "Unit Price",
        "Total Amount",
        "Qty.",
        "Unit Price ($)",
        "Description of Work",
        "Rate per Unit"
    ]
    
    print("Original headers:")
    for header in test_headers:
        print(f"  '{header}'")
    
    normalized = mapper.normalize_header_text(test_headers)
    
    print("\nNormalized headers:")
    for original, normalized_header in zip(test_headers, normalized):
        print(f"  '{original}' -> '{normalized_header}'")


def demo_confidence_calculation():
    """Demonstrate confidence calculation"""
    print("\n=== Confidence Calculation ===")
    
    mapper = ColumnMapper()
    
    # Test different mapping scenarios
    test_cases = [
        {
            "name": "High Confidence",
            "headers": ["Item Code", "Description", "Unit", "Quantity", "Unit Price", "Total Amount"],
            "expected": "High confidence due to clear headers"
        },
        {
            "name": "Medium Confidence",
            "headers": ["Code", "Desc", "Qty", "Rate", "Total"],
            "expected": "Medium confidence due to abbreviated headers"
        },
        {
            "name": "Low Confidence",
            "headers": ["Col1", "Col2", "Col3", "Col4", "Col5"],
            "expected": "Low confidence due to generic headers"
        }
    ]
    
    for test_case in test_cases:
        print(f"\n{test_case['name']}: {test_case['headers']}")
        print("-" * 40)
        
        mappings = mapper.map_columns_to_types(test_case['headers'])
        confidence = mapper.calculate_mapping_confidence(mappings)
        
        print(f"Confidence: {confidence:.2f}")
        print(f"Expected: {test_case['expected']}")
        
        if mappings:
            print("Mappings:")
            for mapping in mappings:
                print(f"  {mapping.column_index + 1}: {mapping.mapped_type.value} ({mapping.confidence:.2f})")


def main():
    """Run all demos"""
    print("Column Mapper Demo")
    print("=" * 50)
    
    demo_header_detection()
    demo_column_mapping()
    demo_complete_mapping()
    demo_alternative_mappings()
    demo_quick_mapping()
    demo_normalization()
    demo_confidence_calculation()
    
    print("\n" + "=" * 50)
    print("All demos completed successfully!")


if __name__ == "__main__":
    main() 