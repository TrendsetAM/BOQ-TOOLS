#!/usr/bin/env python3
"""
Row Classifier Demo
Demonstrates intelligent row classification with data completeness scoring
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.row_classifier import RowClassifier, classify_rows_quick, RowType
from utils.config import ColumnType
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')


def create_sample_sheet_data():
    """Create sample sheet data for testing"""
    return [
        # Header rows
        ["SECTION 1: EARTHWORKS", "", "", "", "", ""],
        ["", "", "", "", "", ""],
        ["1.1", "Excavation for foundation", "m³", "100", "25.50", "2550.00"],
        ["1.2", "Backfilling", "m³", "50", "15.00", "750.00"],
        ["", "Subtotal - Earthworks", "", "", "", "3300.00"],
        ["", "", "", "", "", ""],
        
        # Section break
        ["SECTION 2: CONCRETE WORKS", "", "", "", "", ""],
        ["", "", "", "", "", ""],
        ["2.1", "Concrete foundation", "m³", "50", "150.00", "7500.00"],
        ["2.2", "Concrete columns", "m³", "30", "180.00", "5400.00"],
        ["2.3", "Concrete beams", "m³", "25", "200.00", "5000.00"],
        ["", "Subtotal - Concrete", "", "", "", "17900.00"],
        ["", "", "", "", "", ""],
        
        # Notes and comments
        ["Note: All concrete works include formwork and reinforcement", "", "", "", "", ""],
        ["Note: Prices include 10% contingency", "", "", "", "", ""],
        ["", "", "", "", "", ""],
        
        # More line items
        ["3.1", "Reinforcement steel", "kg", "2000", "2.50", "5000.00"],
        ["3.2", "Formwork", "m²", "200", "15.00", "3000.00"],
        ["", "Subtotal - Steel & Formwork", "", "", "", "8000.00"],
        ["", "", "", "", "", ""],
        
        # Summary rows
        ["", "TOTAL CONSTRUCTION", "", "", "", "29200.00"],
        ["", "Contingency (10%)", "", "", "", "2920.00"],
        ["", "GRAND TOTAL", "", "", "", "32120.00"],
        
        # Invalid line items
        ["4.1", "Invalid item", "", "100", "25.50", "2550.00"],  # Missing unit
        ["4.2", "Negative price", "m³", "50", "-150.00", "-7500.00"],  # Negative price
        ["4.3", "Zero quantity", "m³", "0", "150.00", "0.00"],  # Zero quantity
    ]


def create_column_mapping():
    """Create sample column mapping"""
    return {
        0: ColumnType.CODE,
        1: ColumnType.DESCRIPTION,
        2: ColumnType.UNIT,
        3: ColumnType.QUANTITY,
        4: ColumnType.UNIT_PRICE,
        5: ColumnType.TOTAL_PRICE
    }


def demo_basic_classification():
    """Demonstrate basic row classification"""
    print("=== Basic Row Classification ===")
    
    classifier = RowClassifier()
    sheet_data = create_sample_sheet_data()
    column_mapping = create_column_mapping()
    
    print(f"Classifying {len(sheet_data)} rows...")
    
    result = classifier.classify_rows(sheet_data, column_mapping)
    
    print(f"\nClassification Summary:")
    for row_type, count in result.summary.items():
        if count > 0:
            print(f"  {row_type.value}: {count}")
    
    print(f"\nOverall Quality Score: {result.overall_quality_score:.2f}")
    
    if result.suggestions:
        print("\nSuggestions:")
        for suggestion in result.suggestions:
            print(f"  - {suggestion}")


def demo_detailed_classification():
    """Demonstrate detailed row classification with reasoning"""
    print("\n=== Detailed Row Classification ===")
    
    classifier = RowClassifier()
    sheet_data = create_sample_sheet_data()
    column_mapping = create_column_mapping()
    
    result = classifier.classify_rows(sheet_data, column_mapping)
    
    print("Detailed classifications:")
    print("-" * 80)
    
    for classification in result.classifications:
        print(f"Row {classification.row_index + 1}: {classification.row_type.value}")
        print(f"  Confidence: {classification.confidence:.2f}")
        print(f"  Completeness: {classification.completeness_score:.2f}")
        
        if classification.hierarchical_level is not None:
            print(f"  Hierarchical Level: {classification.hierarchical_level}")
        
        if classification.section_title:
            print(f"  Section Title: '{classification.section_title}'")
        
        if classification.reasoning:
            print(f"  Reasoning: {'; '.join(classification.reasoning)}")
        
        if classification.validation_errors:
            print(f"  Validation Errors: {'; '.join(classification.validation_errors)}")
        
        # Show row data
        row_data = sheet_data[classification.row_index]
        print(f"  Data: {row_data}")
        print()


def demo_completeness_scoring():
    """Demonstrate completeness scoring"""
    print("\n=== Completeness Scoring ===")
    
    classifier = RowClassifier()
    column_mapping = create_column_mapping()
    
    # Test different row types
    test_rows = [
        ["1.1", "Complete item", "m³", "100", "25.50", "2550.00"],  # Complete
        ["1.2", "Missing unit", "", "100", "25.50", "2550.00"],     # Missing unit
        ["1.3", "Missing price", "m³", "100", "", "2550.00"],       # Missing price
        ["1.4", "Only description", "Description only", "", "", ""], # Only description
        ["", "", "", "", "", ""],                                    # Empty
    ]
    
    print("Completeness scores for different row types:")
    print("-" * 60)
    
    for i, row_data in enumerate(test_rows):
        score = classifier.calculate_completeness_score(row_data, column_mapping)
        print(f"Row {i + 1}: {score:.2f} - {row_data}")


def demo_subtotal_detection():
    """Demonstrate subtotal pattern detection"""
    print("\n=== Subtotal Pattern Detection ===")
    
    classifier = RowClassifier()
    
    test_rows = [
        ["", "Subtotal", "", "", "", "1000.00"],
        ["", "Total", "", "", "", "5000.00"],
        ["", "Grand Total", "", "", "", "10000.00"],
        ["", "Sum of items", "", "", "", "2500.00"],
        ["", "Regular item", "m³", "100", "25.50", "2550.00"],
        ["", "Less discount", "", "", "", "-500.00"],
        ["", "Plus tax", "", "", "", "250.00"],
    ]
    
    print("Subtotal detection results:")
    print("-" * 50)
    
    for i, row_data in enumerate(test_rows):
        is_subtotal = classifier.detect_subtotal_patterns(row_data)
        print(f"Row {i + 1}: {'✓' if is_subtotal else '✗'} - {row_data[1] if len(row_data) > 1 else 'N/A'}")


def demo_validation():
    """Demonstrate line item validation"""
    print("\n=== Line Item Validation ===")
    
    classifier = RowClassifier()
    column_mapping = create_column_mapping()
    
    test_rows = [
        ["1.1", "Valid item", "m³", "100", "25.50", "2550.00"],
        ["1.2", "Negative price", "m³", "100", "-25.50", "-2550.00"],
        ["1.3", "Zero quantity", "m³", "0", "25.50", "0.00"],
        ["1.4", "Missing description", "", "100", "25.50", "2550.00"],
        ["1.5", "Invalid quantity", "m³", "abc", "25.50", "2550.00"],
        ["1.6", "Missing unit price", "m³", "100", "", "2550.00"],
    ]
    
    print("Validation results:")
    print("-" * 60)
    
    for i, row_data in enumerate(test_rows):
        errors = classifier.validate_line_item(row_data, column_mapping)
        status = "✓ Valid" if not errors else "✗ Invalid"
        print(f"Row {i + 1}: {status}")
        if errors:
            for error in errors:
                print(f"    Error: {error}")
        print()


def demo_hierarchical_detection():
    """Demonstrate hierarchical numbering detection"""
    print("\n=== Hierarchical Numbering Detection ===")
    
    classifier = RowClassifier()
    
    test_rows = [
        ["1.1", "Item 1.1", "m³", "100", "25.50", "2550.00"],
        ["1.1.1", "Sub-item 1.1.1", "m³", "50", "25.50", "1275.00"],
        ["1.2", "Item 1.2", "m³", "100", "25.50", "2550.00"],
        ["A.1", "Item A.1", "m³", "100", "25.50", "2550.00"],
        ["1-1", "Item 1-1", "m³", "100", "25.50", "2550.00"],
        ["1)1", "Item 1)1", "m³", "100", "25.50", "2550.00"],
        ["Regular", "Regular item", "m³", "100", "25.50", "2550.00"],
    ]
    
    print("Hierarchical level detection:")
    print("-" * 50)
    
    for i, row_data in enumerate(test_rows):
        level = classifier._detect_hierarchical_level(row_data)
        level_str = f"Level {level}" if level is not None else "No hierarchy"
        print(f"Row {i + 1}: {level_str} - {row_data[0] if row_data else 'N/A'}")


def demo_confidence_scoring():
    """Demonstrate confidence scoring"""
    print("\n=== Confidence Scoring ===")
    
    classifier = RowClassifier()
    sheet_data = create_sample_sheet_data()
    column_mapping = create_column_mapping()
    
    result = classifier.classify_rows(sheet_data, column_mapping)
    
    print("Confidence scores by row type:")
    print("-" * 60)
    
    # Group by row type
    by_type = {}
    for classification in result.classifications:
        row_type = classification.row_type
        if row_type not in by_type:
            by_type[row_type] = []
        by_type[row_type].append(classification)
    
    for row_type, classifications in by_type.items():
        avg_confidence = sum(c.confidence for c in classifications) / len(classifications)
        print(f"{row_type.value}: {avg_confidence:.2f} (avg of {len(classifications)} rows)")
        
        # Show range
        min_conf = min(c.confidence for c in classifications)
        max_conf = max(c.confidence for c in classifications)
        print(f"  Range: {min_conf:.2f} - {max_conf:.2f}")


def demo_quick_classification():
    """Demonstrate quick classification function"""
    print("\n=== Quick Classification ===")
    
    sheet_data = create_sample_sheet_data()
    
    # Create string-based column mapping
    column_mapping = {
        0: "code",
        1: "description", 
        2: "unit",
        3: "quantity",
        4: "unit_price",
        5: "total_price"
    }
    
    result = classify_rows_quick(sheet_data, column_mapping)
    
    print("Quick classification results:")
    print("-" * 50)
    
    for row_index, row_type in result.items():
        print(f"Row {row_index + 1}: {row_type}")


def demo_edge_cases():
    """Demonstrate edge cases and error handling"""
    print("\n=== Edge Cases and Error Handling ===")
    
    classifier = RowClassifier()
    column_mapping = create_column_mapping()
    
    edge_cases = [
        [],  # Empty row
        [""],  # Single empty cell
        ["", "", "", "", "", ""],  # All empty cells
        ["1.1", "Item with special chars: @#$%", "m³", "100", "25.50", "2550.00"],  # Special chars
        ["1.2", "Item with currency: $25.50", "m³", "100", "$25.50", "$2,550.00"],  # Currency symbols
        ["1.3", "Item with percentage", "m³", "100", "25.50", "2550.00"],  # Percentage
        ["NOTE: This is a very long note that spans multiple words and should be classified as a comment row", "", "", "", "", ""],  # Long note
    ]
    
    print("Edge case handling:")
    print("-" * 60)
    
    for i, row_data in enumerate(edge_cases):
        try:
            completeness = classifier.calculate_completeness_score(row_data, column_mapping)
            is_subtotal = classifier.detect_subtotal_patterns(row_data)
            errors = classifier.validate_line_item(row_data, column_mapping)
            
            print(f"Case {i + 1}: Completeness={completeness:.2f}, Subtotal={is_subtotal}, Errors={len(errors)}")
            print(f"  Data: {row_data}")
            
        except Exception as e:
            print(f"Case {i + 1}: Error - {e}")
            print(f"  Data: {row_data}")


def main():
    """Run all demos"""
    print("Row Classifier Demo")
    print("=" * 60)
    
    demo_basic_classification()
    demo_detailed_classification()
    demo_completeness_scoring()
    demo_subtotal_detection()
    demo_validation()
    demo_hierarchical_detection()
    demo_confidence_scoring()
    demo_quick_classification()
    demo_edge_cases()
    
    print("\n" + "=" * 60)
    print("All demos completed successfully!")


if __name__ == "__main__":
    main() 