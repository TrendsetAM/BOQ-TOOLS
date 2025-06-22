#!/usr/bin/env python3
"""
Sheet Classifier Demo
Demonstrates the intelligent sheet classification capabilities
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.sheet_classifier import SheetClassifier, SheetType
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')


def create_sample_sheets():
    """Create sample sheet data for testing"""
    sheets = {
        "BOQ Main": {
            "headers": ["Item Code", "Description", "Unit", "Quantity", "Unit Price", "Total Amount"],
            "content": [
                ["001", "Excavation for foundation", "m³", "100", "25.50", "2550.00"],
                ["002", "Concrete foundation", "m³", "50", "150.00", "7500.00"],
                ["003", "Reinforcement steel", "kg", "2000", "2.50", "5000.00"],
                ["004", "Formwork", "m²", "200", "15.00", "3000.00"],
                ["", "", "", "", "", ""],
                ["", "Subtotal", "", "", "", "18050.00"],
                ["", "Contingency (10%)", "", "", "", "1805.00"],
                ["", "Grand Total", "", "", "", "19855.00"]
            ],
            "metadata": {"row_count": 8, "column_count": 6, "data_density": 0.75}
        },
        
        "Summary": {
            "headers": ["Category", "Description", "Amount"],
            "content": [
                ["Substructure", "Foundation works", "18050.00"],
                ["Contingency", "10% contingency", "1805.00"],
                ["", "Total", "19855.00"]
            ],
            "metadata": {"row_count": 3, "column_count": 3, "data_density": 0.89}
        },
        
        "General Information": {
            "headers": ["Project Details", "Value"],
            "content": [
                ["Project Name", "Office Building Project"],
                ["Client", "ABC Construction Ltd"],
                ["Location", "123 Main Street, City"],
                ["Contractor", "XYZ Builders"],
                ["Start Date", "01/01/2024"],
                ["Completion Date", "31/12/2024"],
                ["Project Manager", "John Smith"],
                ["", ""],
                ["Site Conditions", "Urban area, good access"],
                ["Weather Considerations", "Standard construction season"]
            ],
            "metadata": {"row_count": 10, "column_count": 2, "data_density": 0.90}
        },
        
        "Reference Standards": {
            "headers": ["Standard Code", "Description", "Applicable Sections"],
            "content": [
                ["BS 8110", "Structural use of concrete", "All concrete works"],
                ["BS 4449", "Steel for reinforcement", "Reinforcement works"],
                ["BS 6399", "Loading for buildings", "Structural design"],
                ["BS 8000", "Workmanship on building sites", "Quality control"],
                ["BS 8204", "Screeds, bases and inlaid floorings", "Floor finishes"],
                ["BS 5268", "Structural use of timber", "Timber works"],
                ["BS 7671", "Requirements for electrical installations", "Electrical works"]
            ],
            "metadata": {"row_count": 7, "column_count": 3, "data_density": 1.0}
        },
        
        "Mixed Content": {
            "headers": ["Section", "Description", "Quantity", "Notes"],
            "content": [
                ["Preliminaries", "Site setup and temporary works", "1", "Lump sum"],
                ["", "", "", ""],
                ["Substructure", "Foundation and basement works", "", "See detailed BOQ"],
                ["", "Excavation", "100 m³", "Manual excavation"],
                ["", "Concrete", "50 m³", "C25 concrete"],
                ["", "", "", ""],
                ["Superstructure", "Main building structure", "", "See detailed BOQ"],
                ["", "Concrete frame", "200 m³", "C30 concrete"],
                ["", "Steel frame", "50 ton", "Structural steel"],
                ["", "", "", ""],
                ["Summary", "Total project value", "19855", "Including contingency"]
            ],
            "metadata": {"row_count": 11, "column_count": 4, "data_density": 0.68}
        }
    }
    
    return sheets


def demo_basic_classification():
    """Demonstrate basic sheet classification"""
    print("=== Basic Sheet Classification ===")
    
    classifier = SheetClassifier()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nClassifying: {sheet_name}")
        print("-" * 40)
        
        result = classifier.classify_sheet(sheet_data, sheet_name)
        
        print(f"Type: {result.sheet_type.value}")
        print(f"Confidence: {result.confidence:.2f}")
        print(f"Scores: Keyword={result.scores['keyword']:.2f}, "
              f"Numeric={result.scores['numeric']:.2f}, "
              f"Pattern={result.scores['pattern']:.2f}")
        
        print("Reasoning:")
        for reason in result.reasoning[:5]:  # Show first 5 reasons
            print(f"  {reason}")
        
        if result.keyword_matches:
            print("Keyword matches:")
            for match in result.keyword_matches[:3]:  # Show first 3 matches
                print(f"  - {match}")
        
        if result.patterns_detected:
            print("Patterns detected:")
            for pattern in result.patterns_detected[:3]:  # Show first 3 patterns
                print(f"  - {pattern}")


def demo_numeric_analysis():
    """Demonstrate numeric content analysis"""
    print("\n=== Numeric Content Analysis ===")
    
    classifier = SheetClassifier()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nAnalyzing: {sheet_name}")
        print("-" * 30)
        
        content = sheet_data['content']
        headers = sheet_data['headers']
        
        numeric_result = classifier.calculate_numeric_ratio(content, headers)
        
        print(f"Numeric ratio: {numeric_result['ratio']:.2f}")
        print(f"Analysis: {numeric_result['analysis']}")
        print(f"Numeric columns: {numeric_result['numeric_columns']}")


def demo_pattern_detection():
    """Demonstrate pattern detection"""
    print("\n=== Pattern Detection ===")
    
    classifier = SheetClassifier()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nDetecting patterns: {sheet_name}")
        print("-" * 35)
        
        content = sheet_data['content']
        headers = sheet_data['headers']
        
        pattern_result = classifier.detect_patterns(content, headers)
        
        print(f"Pattern score: {pattern_result['score']:.2f}")
        if pattern_result['patterns']:
            print("Detected patterns:")
            for pattern in pattern_result['patterns']:
                print(f"  - {pattern}")
        else:
            print("  No significant patterns detected")


def demo_keyword_scoring():
    """Demonstrate keyword scoring"""
    print("\n=== Keyword Scoring ===")
    
    classifier = SheetClassifier()
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        print(f"\nScoring keywords: {sheet_name}")
        print("-" * 30)
        
        content = sheet_data['content']
        headers = sheet_data['headers']
        
        keyword_result = classifier.score_keywords(sheet_name, content, headers)
        
        print(f"Keyword score: {keyword_result['score']:.2f}")
        if keyword_result['matches']:
            print("Matches:")
            for match in keyword_result['matches']:
                print(f"  - {match}")
        else:
            print("  No keyword matches")


def demo_classification_summary():
    """Demonstrate classification summary"""
    print("\n=== Classification Summary ===")
    
    classifier = SheetClassifier()
    sheets = create_sample_sheets()
    
    results = []
    for sheet_name, sheet_data in sheets.items():
        result = classifier.classify_sheet(sheet_data, sheet_name)
        results.append(result)
    
    summary = classifier.get_classification_summary(results)
    
    print(f"Total sheets analyzed: {summary['total_sheets']}")
    print(f"Confidence stats: Avg={summary['confidence_stats']['average']:.2f}, "
          f"Min={summary['confidence_stats']['min']:.2f}, "
          f"Max={summary['confidence_stats']['max']:.2f}")
    
    print("\nSheet type distribution:")
    for sheet_type, count in summary['sheet_types'].items():
        percentage = summary['type_distribution'][sheet_type] * 100
        print(f"  {sheet_type}: {count} sheets ({percentage:.1f}%)")


def demo_quick_classification():
    """Demonstrate quick classification function"""
    print("\n=== Quick Classification ===")
    
    from core.sheet_classifier import classify_sheet_quick
    
    sheets = create_sample_sheets()
    
    for sheet_name, sheet_data in sheets.items():
        sheet_type, confidence = classify_sheet_quick(sheet_data, sheet_name)
        print(f"{sheet_name}: {sheet_type} (confidence: {confidence:.2f})")


def main():
    """Run all demos"""
    print("Sheet Classifier Demo")
    print("=" * 50)
    
    demo_basic_classification()
    demo_numeric_analysis()
    demo_pattern_detection()
    demo_keyword_scoring()
    demo_classification_summary()
    demo_quick_classification()
    
    print("\n" + "=" * 50)
    print("All demos completed successfully!")


if __name__ == "__main__":
    main() 