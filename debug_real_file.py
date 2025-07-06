#!/usr/bin/env python3

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.file_processor import ExcelProcessor
from core.column_mapper import ColumnMapper
from pathlib import Path

def debug_real_excel_file():
    print("=== Debug Real Excel File Structure ===")
    
    # The file from the screenshot appears to be processed, let's check what we can see
    print("\n1. Testing with a simulated structure based on the UI screenshot:")
    
    # Based on the UI screenshot, it looks like there might be more columns
    # Let's simulate what the actual Excel structure might look like
    simulated_data = [
        # This might be what the actual Excel file looks like
        ['Code', 'DESCRIPTION', 'SCOPE', 'Unit of measure', 'Quantity', 'Labour', '', '', 'Equipment', '', '', '', 'SUPPLY', ''],
        ['001', 'Paint work', 'Mandatory', 'sqm', '100', '50', '', '', '25', '', '', '', '75', ''],
        ['002', 'Concrete work', 'Optional', 'cum', '20', '80', '', '', '40', '', '', '', '60', '']
    ]
    
    cm = ColumnMapper()
    
    print("Simulated raw data:")
    for i, row in enumerate(simulated_data):
        print(f"  Row {i}: {row}")
    
    print(f"\nNumber of columns: {len(simulated_data[0])}")
    
    result = cm.process_sheet_mapping(simulated_data)
    print(f"\nHeader row detected: {result.header_row.row_index}")
    print(f"Is merged: {result.header_row.is_merged}")
    print(f"Headers found: {result.header_row.headers}")
    print(f"Number of headers: {len(result.header_row.headers)}")
    
    print("\nDetailed column mappings:")
    for i, mapping in enumerate(result.mappings):
        header = mapping.original_header if mapping.original_header else "(empty)"
        print(f"  Col {mapping.column_index:2d}: '{header}' -> {mapping.mapped_type} (conf: {mapping.confidence:.2f})")
    
    print(f"\nTotal mappings: {len(result.mappings)}")
    print(f"Overall confidence: {result.overall_confidence:.2f}")
    
    # Also test with a cleaner structure (no empty columns)
    print("\n" + "="*60)
    print("2. Testing with clean structure (no empty columns):")
    
    clean_data = [
        ['Code', 'DESCRIPTION', 'SCOPE', 'Unit of measure', 'Quantity', 'Labour', 'Equipment', 'SUPPLY'],
        ['001', 'Paint work', 'Mandatory', 'sqm', '100', '50', '25', '75'],
        ['002', 'Concrete work', 'Optional', 'cum', '20', '80', '40', '60']
    ]
    
    print("Clean raw data:")
    for i, row in enumerate(clean_data):
        print(f"  Row {i}: {row}")
    
    clean_result = cm.process_sheet_mapping(clean_data)
    print(f"\nHeader row detected: {clean_result.header_row.row_index}")
    print(f"Is merged: {clean_result.header_row.is_merged}")
    print(f"Headers found: {clean_result.header_row.headers}")
    
    print("\nClean column mappings:")
    for i, mapping in enumerate(clean_result.mappings):
        print(f"  Col {mapping.column_index}: '{mapping.original_header}' -> {mapping.mapped_type} (conf: {mapping.confidence:.2f})")
    
    print(f"\nClean overall confidence: {clean_result.overall_confidence:.2f}")

if __name__ == "__main__":
    debug_real_excel_file() 