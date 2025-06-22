"""
Mapping Generator Demo
Demonstrates the unified mapping structure generation capabilities
"""

import sys
import os
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.mapping_generator import MappingGenerator, ProcessingStatus, ReviewFlag, generate_mapping_quick
from utils.config import get_config
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_sample_processor_results():
    """Create sample processor results for testing"""
    
    # Sample file info
    file_info = {
        'filename': 'sample_boq.xlsx',
        'file_path': '/path/to/sample_boq.xlsx',
        'file_size_mb': 2.5,
        'file_format': 'xlsx',
        'total_sheets': 3,
        'visible_sheets': 3
    }
    
    # Sample sheet data
    sheet_data = {
        'Sheet1': [
            ['Item No.', 'Description', 'Unit', 'Quantity', 'Unit Price', 'Total Price'],
            ['1.1', 'Excavation for foundation', 'm³', '150.00', '25.00', '3750.00'],
            ['1.2', 'Concrete foundation', 'm³', '75.00', '120.00', '9000.00'],
            ['', '', '', '', 'Subtotal', '12750.00'],
            ['2.1', 'Brickwork', 'm²', '200.00', '45.00', '9000.00'],
            ['', '', '', '', 'Total', '21750.00']
        ],
        'Sheet2': [
            ['Project Information', '', '', '', '', ''],
            ['Project Name:', 'Sample Building Project', '', '', '', ''],
            ['Client:', 'ABC Construction Ltd', '', '', '', ''],
            ['Date:', '2024-01-15', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['Item', 'Specification', 'Qty', 'Rate', 'Amount', 'Remarks'],
            ['1', 'Site preparation', '1', '5000.00', '5000.00', ''],
            ['2', 'Foundation work', '1', '15000.00', '15000.00', ''],
            ['', '', '', '', 'Total', '20000.00']
        ],
        'Sheet3': [
            ['Notes and Conditions', '', '', '', '', ''],
            ['1. All prices are inclusive of taxes', '', '', '', '', ''],
            ['2. Payment terms: 30 days', '', '', '', '', ''],
            ['3. Delivery: As per schedule', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['Contact Information', '', '', '', '', ''],
            ['Phone:', '+1-234-567-8900', '', '', '', ''],
            ['Email:', 'info@sample.com', '', '', '', '']
        ]
    }
    
    # Sample column mappings
    column_mappings = {
        'Sheet1': {
            'header_row_index': 0,
            'header_confidence': 0.95,
            'overall_confidence': 0.88,
            'mappings': [
                {
                    'column_index': 0,
                    'original_header': 'Item No.',
                    'normalized_header': 'item_no',
                    'mapped_type': 'item_number',
                    'confidence': 0.95,
                    'alternatives': [{'type': 'item_number', 'confidence': 0.95}],
                    'reasoning': ['Exact match with item_number keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 1,
                    'original_header': 'Description',
                    'normalized_header': 'description',
                    'mapped_type': 'description',
                    'confidence': 0.90,
                    'alternatives': [{'type': 'description', 'confidence': 0.90}],
                    'reasoning': ['Exact match with description keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 2,
                    'original_header': 'Unit',
                    'normalized_header': 'unit',
                    'mapped_type': 'unit',
                    'confidence': 0.85,
                    'alternatives': [{'type': 'unit', 'confidence': 0.85}],
                    'reasoning': ['Exact match with unit keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 3,
                    'original_header': 'Quantity',
                    'normalized_header': 'quantity',
                    'mapped_type': 'quantity',
                    'confidence': 0.92,
                    'alternatives': [{'type': 'quantity', 'confidence': 0.92}],
                    'reasoning': ['Exact match with quantity keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 4,
                    'original_header': 'Unit Price',
                    'normalized_header': 'unit_price',
                    'mapped_type': 'unit_price',
                    'confidence': 0.88,
                    'alternatives': [{'type': 'unit_price', 'confidence': 0.88}],
                    'reasoning': ['Exact match with unit_price keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 5,
                    'original_header': 'Total Price',
                    'normalized_header': 'total_price',
                    'mapped_type': 'total_price',
                    'confidence': 0.90,
                    'alternatives': [{'type': 'total_price', 'confidence': 0.90}],
                    'reasoning': ['Exact match with total_price keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                }
            ],
            'unmapped_columns': [],
            'suggestions': ['All columns mapped successfully']
        },
        'Sheet2': {
            'header_row_index': 5,
            'header_confidence': 0.75,
            'overall_confidence': 0.70,
            'mappings': [
                {
                    'column_index': 0,
                    'original_header': 'Item',
                    'normalized_header': 'item',
                    'mapped_type': 'item_number',
                    'confidence': 0.65,
                    'alternatives': [
                        {'type': 'item_number', 'confidence': 0.65},
                        {'type': 'description', 'confidence': 0.35}
                    ],
                    'reasoning': ['Partial match with item_number keywords'],
                    'is_required': True,
                    'validation_status': 'needs_review'
                },
                {
                    'column_index': 1,
                    'original_header': 'Specification',
                    'normalized_header': 'specification',
                    'mapped_type': 'description',
                    'confidence': 0.80,
                    'alternatives': [{'type': 'description', 'confidence': 0.80}],
                    'reasoning': ['Close match with description keywords'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 2,
                    'original_header': 'Qty',
                    'normalized_header': 'qty',
                    'mapped_type': 'quantity',
                    'confidence': 0.85,
                    'alternatives': [{'type': 'quantity', 'confidence': 0.85}],
                    'reasoning': ['Abbreviation of quantity'],
                    'is_required': True,
                    'validation_status': 'valid'
                },
                {
                    'column_index': 3,
                    'original_header': 'Rate',
                    'normalized_header': 'rate',
                    'mapped_type': 'unit_price',
                    'confidence': 0.70,
                    'alternatives': [
                        {'type': 'unit_price', 'confidence': 0.70},
                        {'type': 'rate', 'confidence': 0.30}
                    ],
                    'reasoning': ['Partial match with unit_price keywords'],
                    'is_required': True,
                    'validation_status': 'needs_review'
                },
                {
                    'column_index': 4,
                    'original_header': 'Amount',
                    'normalized_header': 'amount',
                    'mapped_type': 'total_price',
                    'confidence': 0.75,
                    'alternatives': [
                        {'type': 'total_price', 'confidence': 0.75},
                        {'type': 'amount', 'confidence': 0.25}
                    ],
                    'reasoning': ['Partial match with total_price keywords'],
                    'is_required': True,
                    'validation_status': 'needs_review'
                },
                {
                    'column_index': 5,
                    'original_header': 'Remarks',
                    'normalized_header': 'remarks',
                    'mapped_type': 'notes',
                    'confidence': 0.60,
                    'alternatives': [{'type': 'notes', 'confidence': 0.60}],
                    'reasoning': ['Best match with notes keywords'],
                    'is_required': False,
                    'validation_status': 'valid'
                }
            ],
            'unmapped_columns': [],
            'suggestions': ['Some column mappings have low confidence and need review']
        },
        'Sheet3': {
            'header_row_index': 0,
            'header_confidence': 0.30,
            'overall_confidence': 0.25,
            'mappings': [
                {
                    'column_index': 0,
                    'original_header': 'Notes and Conditions',
                    'normalized_header': 'notes_and_conditions',
                    'mapped_type': 'notes',
                    'confidence': 0.40,
                    'alternatives': [{'type': 'notes', 'confidence': 0.40}],
                    'reasoning': ['Best match with notes keywords'],
                    'is_required': False,
                    'validation_status': 'valid'
                }
            ],
            'unmapped_columns': [1, 2, 3, 4, 5],
            'suggestions': ['This sheet appears to be reference information, not main BOQ data']
        }
    }
    
    # Sample row classifications
    row_classifications = {
        'Sheet1': {
            'overall_quality_score': 0.85,
            'classifications': [
                {
                    'row_index': 0,
                    'row_type': 'header',
                    'confidence': 0.95,
                    'completeness_score': 1.0,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contains all required column headers']
                },
                {
                    'row_index': 1,
                    'row_type': 'line_item',
                    'confidence': 0.90,
                    'completeness_score': 0.95,
                    'hierarchical_level': 1,
                    'section_title': 'Excavation',
                    'validation_errors': [],
                    'reasoning': ['Complete line item with all required fields']
                },
                {
                    'row_index': 2,
                    'row_type': 'line_item',
                    'confidence': 0.90,
                    'completeness_score': 0.95,
                    'hierarchical_level': 1,
                    'section_title': 'Concrete',
                    'validation_errors': [],
                    'reasoning': ['Complete line item with all required fields']
                },
                {
                    'row_index': 3,
                    'row_type': 'subtotal',
                    'confidence': 0.85,
                    'completeness_score': 0.30,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contains subtotal calculation']
                },
                {
                    'row_index': 4,
                    'row_type': 'line_item',
                    'confidence': 0.90,
                    'completeness_score': 0.95,
                    'hierarchical_level': 2,
                    'section_title': 'Brickwork',
                    'validation_errors': [],
                    'reasoning': ['Complete line item with all required fields']
                },
                {
                    'row_index': 5,
                    'row_type': 'total',
                    'confidence': 0.85,
                    'completeness_score': 0.30,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contains total calculation']
                }
            ],
            'suggestions': ['Well-structured BOQ with clear hierarchy']
        },
        'Sheet2': {
            'overall_quality_score': 0.70,
            'classifications': [
                {
                    'row_index': 0,
                    'row_type': 'header',
                    'confidence': 0.75,
                    'completeness_score': 0.80,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Project information header']
                },
                {
                    'row_index': 1,
                    'row_type': 'info',
                    'confidence': 0.80,
                    'completeness_score': 0.60,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Project name information']
                },
                {
                    'row_index': 2,
                    'row_type': 'info',
                    'confidence': 0.80,
                    'completeness_score': 0.60,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Client information']
                },
                {
                    'row_index': 3,
                    'row_type': 'info',
                    'confidence': 0.80,
                    'completeness_score': 0.60,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Date information']
                },
                {
                    'row_index': 4,
                    'row_type': 'blank',
                    'confidence': 0.90,
                    'completeness_score': 0.0,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Empty row']
                },
                {
                    'row_index': 5,
                    'row_type': 'header',
                    'confidence': 0.75,
                    'completeness_score': 0.80,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Data table header']
                },
                {
                    'row_index': 6,
                    'row_type': 'line_item',
                    'confidence': 0.70,
                    'completeness_score': 0.85,
                    'hierarchical_level': 1,
                    'section_title': 'Site Preparation',
                    'validation_errors': ['Missing unit information'],
                    'reasoning': ['Line item with some missing data']
                },
                {
                    'row_index': 7,
                    'row_type': 'line_item',
                    'confidence': 0.70,
                    'completeness_score': 0.85,
                    'hierarchical_level': 1,
                    'section_title': 'Foundation',
                    'validation_errors': ['Missing unit information'],
                    'reasoning': ['Line item with some missing data']
                },
                {
                    'row_index': 8,
                    'row_type': 'total',
                    'confidence': 0.85,
                    'completeness_score': 0.30,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contains total calculation']
                }
            ],
            'suggestions': ['Some line items missing unit information']
        },
        'Sheet3': {
            'overall_quality_score': 0.30,
            'classifications': [
                {
                    'row_index': 0,
                    'row_type': 'header',
                    'confidence': 0.40,
                    'completeness_score': 0.20,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Reference information header']
                },
                {
                    'row_index': 1,
                    'row_type': 'notes',
                    'confidence': 0.60,
                    'completeness_score': 0.20,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Condition note']
                },
                {
                    'row_index': 2,
                    'row_type': 'notes',
                    'confidence': 0.60,
                    'completeness_score': 0.20,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Condition note']
                },
                {
                    'row_index': 3,
                    'row_type': 'notes',
                    'confidence': 0.60,
                    'completeness_score': 0.20,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Condition note']
                },
                {
                    'row_index': 4,
                    'row_type': 'blank',
                    'confidence': 0.90,
                    'completeness_score': 0.0,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Empty row']
                },
                {
                    'row_index': 5,
                    'row_type': 'header',
                    'confidence': 0.40,
                    'completeness_score': 0.20,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contact information header']
                },
                {
                    'row_index': 6,
                    'row_type': 'info',
                    'confidence': 0.50,
                    'completeness_score': 0.30,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contact information']
                },
                {
                    'row_index': 7,
                    'row_type': 'info',
                    'confidence': 0.50,
                    'completeness_score': 0.30,
                    'hierarchical_level': None,
                    'section_title': None,
                    'validation_errors': [],
                    'reasoning': ['Contact information']
                }
            ],
            'suggestions': ['This sheet contains reference information, not BOQ data']
        }
    }
    
    # Sample validation results
    validation_results = {
        'Sheet1': {
            'overall_score': 0.92,
            'mathematical_consistency': 0.95,
            'data_type_quality': 0.90,
            'business_rule_compliance': 0.90,
            'error_count': 0,
            'warning_count': 1,
            'info_count': 2,
            'suggestions': ['Mathematical calculations are consistent', 'Data types are appropriate']
        },
        'Sheet2': {
            'overall_score': 0.75,
            'mathematical_consistency': 0.80,
            'data_type_quality': 0.70,
            'business_rule_compliance': 0.75,
            'error_count': 2,
            'warning_count': 3,
            'info_count': 1,
            'suggestions': ['Some missing unit information', 'Consider adding more detailed specifications']
        },
        'Sheet3': {
            'overall_score': 0.40,
            'mathematical_consistency': 0.50,
            'data_type_quality': 0.30,
            'business_rule_compliance': 0.40,
            'error_count': 0,
            'warning_count': 0,
            'info_count': 5,
            'suggestions': ['This sheet contains reference information only']
        }
    }
    
    return {
        'file_info': file_info,
        'sheet_data': sheet_data,
        'column_mappings': column_mappings,
        'row_classifications': row_classifications,
        'validation_results': validation_results
    }


def demo_mapping_generator():
    """Demonstrate the MappingGenerator functionality"""
    
    print("=" * 60)
    print("MAPPING GENERATOR DEMO")
    print("=" * 60)
    
    # Create sample data
    print("\n1. Creating sample processor results...")
    processor_results = create_sample_processor_results()
    
    # Initialize mapping generator
    print("\n2. Initializing MappingGenerator...")
    generator = MappingGenerator(processing_version="1.0.0")
    
    # Generate file mapping
    print("\n3. Generating unified file mapping...")
    file_mapping = generator.generate_file_mapping(processor_results)
    
    # Display results
    print("\n4. File Mapping Results:")
    print(f"   - Filename: {file_mapping.metadata.filename}")
    print(f"   - Total Sheets: {file_mapping.metadata.total_sheets}")
    print(f"   - Global Confidence: {file_mapping.global_confidence:.2f}")
    print(f"   - Export Ready: {file_mapping.export_ready}")
    print(f"   - Review Flags: {[flag.value for flag in file_mapping.review_flags]}")
    
    # Display processing summary
    print("\n5. Processing Summary:")
    summary = file_mapping.processing_summary
    print(f"   - Total Rows Processed: {summary.total_rows_processed}")
    print(f"   - Total Columns Mapped: {summary.total_columns_mapped}")
    print(f"   - Successful Sheets: {summary.successful_sheets}")
    print(f"   - Partial Sheets: {summary.partial_sheets}")
    print(f"   - Failed Sheets: {summary.failed_sheets}")
    print(f"   - Sheets Needing Review: {summary.sheets_needing_review}")
    print(f"   - Average Confidence: {summary.average_confidence:.2f}")
    print(f"   - Total Validation Errors: {summary.total_validation_errors}")
    print(f"   - Total Validation Warnings: {summary.total_validation_warnings}")
    
    # Display recommendations
    print("\n6. Recommendations:")
    for i, rec in enumerate(summary.recommendations, 1):
        print(f"   {i}. {rec}")
    
    # Display sheet details
    print("\n7. Sheet Details:")
    for i, sheet in enumerate(file_mapping.sheets, 1):
        print(f"\n   Sheet {i}: {sheet.sheet_name}")
        print(f"   - Status: {sheet.processing_status.value}")
        print(f"   - Rows: {sheet.row_count}, Columns: {sheet.column_count}")
        print(f"   - Overall Confidence: {sheet.overall_confidence:.2f}")
        print(f"   - Review Flags: {[flag.value for flag in sheet.review_flags]}")
        
        if sheet.manual_review_items:
            print(f"   - Manual Review Items: {len(sheet.manual_review_items)}")
        
        if sheet.warnings:
            print(f"   - Warnings: {len(sheet.warnings)}")
    
    # Test manual review flagging
    print("\n8. Manual Review Items:")
    review_items = generator.flag_manual_review_items(file_mapping.sheets)
    for i, item in enumerate(review_items, 1):
        print(f"   {i}. {item['type']} - {item.get('suggestion', 'No suggestion')}")
    
    # Test JSON export
    print("\n9. Testing JSON Export...")
    try:
        json_output = generator.export_mapping_to_json(file_mapping)
        print(f"   - JSON export successful ({len(json_output)} characters)")
        
        # Save to file for inspection
        output_file = project_root / "examples" / "mapping_export.json"
        generator.export_mapping_to_json(file_mapping, output_file)
        print(f"   - Saved to: {output_file}")
        
    except Exception as e:
        print(f"   - JSON export failed: {e}")
    
    # Test quick mapping
    print("\n10. Quick Mapping Test:")
    quick_result = generate_mapping_quick(processor_results)
    print(f"   - Global Confidence: {quick_result['global_confidence']:.2f}")
    print(f"   - Sheet Count: {quick_result['sheet_count']}")
    print(f"   - Export Ready: {quick_result['export_ready']}")
    print(f"   - Review Flags: {quick_result['review_flags']}")
    
    print("\n" + "=" * 60)
    print("DEMO COMPLETED SUCCESSFULLY!")
    print("=" * 60)


if __name__ == "__main__":
    demo_mapping_generator() 