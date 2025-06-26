"""
Manual Categorizer for BOQ Tools
Generates Excel files for manual categorization of unmatched descriptions
"""

import logging
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

from core.auto_categorizer import UnmatchedDescription

logger = logging.getLogger(__name__)


def generate_manual_categorization_excel(unmatched_descriptions: List[UnmatchedDescription],
                                        available_categories: List[str],
                                        output_dir: Optional[Path] = None) -> Path:
    """
    Generate Excel file for manual categorization of unmatched descriptions
    
    Args:
        unmatched_descriptions: List of UnmatchedDescription objects
        available_categories: List of available categories for dropdown
        output_dir: Output directory (default: current directory)
        
    Returns:
        Path to the created Excel file
    """
    logger.info(f"Generating manual categorization Excel file for {len(unmatched_descriptions)} descriptions")
    
    # Create output directory if it doesn't exist
    if output_dir is None:
        output_dir = Path.cwd()
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"manual_categorization_{timestamp}.xlsx"
    filepath = output_dir / filename
    
    # Create workbook
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    default_sheet = wb.active
    if default_sheet:
        wb.remove(default_sheet)
    
    # Create main categorization sheet
    ws_categorize = wb.create_sheet("Categorization", 0)
    
    # Create instructions sheet
    ws_instructions = wb.create_sheet("Instructions", 1)
    
    # Set up the main categorization sheet
    _setup_categorization_sheet(ws_categorize, unmatched_descriptions, available_categories)
    
    # Set up the instructions sheet
    _setup_instructions_sheet(ws_instructions, available_categories)
    
    # Save the workbook
    wb.save(filepath)
    logger.info(f"Manual categorization Excel file created: {filepath}")
    
    return filepath


def _setup_categorization_sheet(worksheet, unmatched_descriptions: List[UnmatchedDescription], 
                               available_categories: List[str]):
    """Set up the main categorization worksheet"""
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Set up headers
    headers = ["Description", "Source_Sheet", "Frequency", "Category", "Notes"]
    for col, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Add data
    for row, desc in enumerate(unmatched_descriptions, 2):
        worksheet.cell(row=row, column=1, value=desc.description)
        worksheet.cell(row=row, column=2, value=desc.source_sheet_name)
        worksheet.cell(row=row, column=3, value=desc.frequency)
        worksheet.cell(row=row, column=4, value="")  # Category (to be filled manually)
        worksheet.cell(row=row, column=5, value="")  # Notes (to be filled manually)
        
        # Apply borders to all cells
        for col in range(1, 6):
            worksheet.cell(row=row, column=col).border = border
    
    # Set up data validation for Category column
    category_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(available_categories)}"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid Category",
        error="Please select a category from the dropdown list.",
        showInputMessage=True,
        promptTitle="Category Selection",
        prompt="Select a category from the dropdown list."
    )
    worksheet.add_data_validation(category_validation)
    
    # Apply validation to Category column (column D)
    category_validation.add(f'D2:D{len(unmatched_descriptions) + 1}')
    
    # Set column widths
    worksheet.column_dimensions['A'].width = 60  # Description
    worksheet.column_dimensions['B'].width = 20  # Source_Sheet
    worksheet.column_dimensions['C'].width = 12  # Frequency
    worksheet.column_dimensions['D'].width = 25  # Category
    worksheet.column_dimensions['E'].width = 30  # Notes
    
    # Add conditional formatting for frequency
    from openpyxl.formatting.rule import ColorScaleRule
    frequency_rule = ColorScaleRule(
        start_type='min',
        start_color='FFFFFF',
        end_type='max',
        end_color='FF6B6B'
    )
    worksheet.conditional_formatting.add(f'C2:C{len(unmatched_descriptions) + 1}', frequency_rule)
    
    # Freeze the header row
    worksheet.freeze_panes = "A2"


def _setup_instructions_sheet(worksheet, available_categories: List[str]):
    """Set up the instructions worksheet"""
    
    # Title
    title_cell = worksheet['A1']
    title_cell.value = "Manual Categorization Instructions"
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center")
    
    # Instructions
    instructions = [
        ("Purpose", "This file contains descriptions that could not be automatically categorized. Please assign appropriate categories to each description."),
        ("", ""),
        ("How to Use", "1. Go to the 'Categorization' sheet"),
        ("", "2. For each description, select a category from the dropdown in the 'Category' column"),
        ("", "3. Optionally add notes in the 'Notes' column"),
        ("", "4. Save the file when complete"),
        ("", ""),
        ("Available Categories", f"The following categories are available: {', '.join(available_categories)}"),
        ("", ""),
        ("Tips", "- Descriptions are sorted by frequency (most frequent first)"),
        ("", "- Use the 'Notes' column to record any special considerations"),
        ("", "- If a description doesn't fit any category, leave it blank"),
        ("", "- You can add new categories by editing the dropdown list"),
        ("", ""),
        ("File Information", f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"),
        ("", f"Total descriptions to categorize: {len(available_categories)}")
    ]
    
    # Add instructions to worksheet
    for row, (title, content) in enumerate(instructions, 3):
        if title:
            cell = worksheet.cell(row=row, column=1, value=title)
            cell.font = Font(bold=True)
        if content:
            worksheet.cell(row=row, column=2, value=content)
    
    # Set column widths
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 80


def load_manual_categorization_results(filepath: Path) -> List[dict]:
    """
    Load manual categorization results from Excel file
    
    Args:
        filepath: Path to the Excel file with manual categorizations
        
    Returns:
        List of dictionaries with categorization results
    """
    logger.info(f"Loading manual categorization results from {filepath}")
    
    if not filepath.exists():
        raise FileNotFoundError(f"Manual categorization file not found: {filepath}")
    
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Categorization"]
        
        results = []
        
        # Read data starting from row 2 (skip header)
        for row in range(2, ws.max_row + 1):
            description = ws.cell(row=row, column=1).value
            source_sheet = ws.cell(row=row, column=2).value
            frequency = ws.cell(row=row, column=3).value
            category = ws.cell(row=row, column=4).value
            notes = ws.cell(row=row, column=5).value
            
            if description and category:  # Only include rows with both description and category
                # Safely convert frequency to int
                freq_value = 1
                if frequency is not None:
                    try:
                        freq_value = int(float(str(frequency)))
                    except (ValueError, TypeError):
                        freq_value = 1
                
                results.append({
                    'description': str(description).strip(),
                    'source_sheet': str(source_sheet).strip() if source_sheet else 'Unknown',
                    'frequency': freq_value,
                    'category': str(category).strip(),
                    'notes': str(notes).strip() if notes else ''
                })
        
        logger.info(f"Loaded {len(results)} manual categorizations")
        return results
        
    except Exception as e:
        logger.error(f"Error loading manual categorization results: {e}")
        raise


def apply_manual_categorizations(dataframe, manual_results: List[dict], 
                                description_column: str = 'Description',
                                category_column: str = 'Category'):
    """
    Apply manual categorizations to the DataFrame
    
    Args:
        dataframe: DataFrame to update
        manual_results: List of manual categorization results
        description_column: Name of the description column
        category_column: Name of the category column
        
    Returns:
        Updated DataFrame with manual categorizations applied
    """
    logger.info(f"Applying {len(manual_results)} manual categorizations to DataFrame")
    
    # Create a copy to avoid modifying the original
    df = dataframe.copy()
    
    # Create a mapping from description to category
    manual_mapping = {}
    for result in manual_results:
        description = result['description'].lower()
        category = result['category']
        manual_mapping[description] = category
    
    # Apply manual categorizations
    updated_count = 0
    for index, row in df.iterrows():
        description = str(row[description_column]).strip().lower()
        
        if description in manual_mapping:
            df.at[index, category_column] = manual_mapping[description]
            updated_count += 1
    
    logger.info(f"Updated {updated_count} rows with manual categorizations")
    return df


def create_categorization_summary(dataframe, category_column: str = 'Category') -> dict:
    """
    Create a summary of categorization results
    
    Args:
        dataframe: Categorized DataFrame
        category_column: Name of the category column
        
    Returns:
        Dictionary with categorization summary
    """
    total_rows = len(dataframe)
    categorized_rows = len(dataframe[dataframe[category_column].notna() & (dataframe[category_column] != '')])
    uncategorized_rows = total_rows - categorized_rows
    
    category_distribution = dataframe[category_column].value_counts().to_dict()
    
    # Remove empty categories from distribution
    if '' in category_distribution:
        del category_distribution['']
    
    summary = {
        'total_rows': total_rows,
        'categorized_rows': categorized_rows,
        'uncategorized_rows': uncategorized_rows,
        'categorization_rate': categorized_rows / total_rows if total_rows > 0 else 0.0,
        'category_distribution': category_distribution,
        'unique_categories': len(category_distribution)
    }
    
    return summary 


def process_manual_categorizations(excel_filepath: Path, 
                                  description_column: str = "Description",
                                  category_column: str = "Category",
                                  source_sheet_column: str = "Source_Sheet",
                                  notes_column: str = "Notes") -> Dict[str, str]:
    """
    Process user-completed manual categorization Excel file
    
    Args:
        excel_filepath: Path to the completed Excel file
        description_column: Name of the description column
        category_column: Name of the category column
        source_sheet_column: Name of the source sheet column
        notes_column: Name of the notes column
        
    Returns:
        Dictionary mapping descriptions to their manually assigned categories
        
    Raises:
        FileNotFoundError: If the Excel file doesn't exist
        ValueError: If the file is corrupted or missing required columns
        KeyError: If required data is missing
    """
    logger.info(f"Processing manual categorizations from {excel_filepath}")
    
    # Validate file exists
    if not excel_filepath.exists():
        raise FileNotFoundError(f"Manual categorization file not found: {excel_filepath}")
    
    try:
        # Read the Excel file using pandas
        df: pd.DataFrame = pd.read_excel(excel_filepath, sheet_name="Categorization")
        logger.info(f"Successfully loaded Excel file with {len(df)} rows")
        
    except Exception as e:
        logger.error(f"Failed to read Excel file: {e}")
        raise ValueError(f"Failed to read Excel file {excel_filepath}: {e}")
    
    # Validate required columns exist
    required_columns = [description_column, category_column, source_sheet_column]
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        available_columns = list(df.columns)
        logger.error(f"Missing required columns: {missing_columns}")
        logger.error(f"Available columns: {available_columns}")
        raise ValueError(f"Missing required columns: {missing_columns}. Available columns: {available_columns}")
    
    # Clean and validate data
    logger.info("Cleaning and validating data...")
    
    # Remove rows where description is empty or null
    initial_count = len(df)
    df = df.dropna(subset=[description_column])
    df = df[df[description_column].astype(str).str.strip() != '']
    
    if len(df) < initial_count:
        logger.warning(f"Removed {initial_count - len(df)} rows with empty descriptions")
    
    # Clean description column
    df[description_column] = df[description_column].astype(str).str.strip()
    
    # Clean category column and filter for rows with categories
    df[category_column] = df[category_column].astype(str).str.strip()
    df = df[df[category_column] != '']
    df = df[df[category_column] != 'nan']
    
    if len(df) == 0:
        logger.warning("No valid categorizations found in the file")
        return {}
    
    # Clean source sheet column
    if source_sheet_column in df.columns:
        df[source_sheet_column] = df[source_sheet_column].astype(str).str.strip()
        df[source_sheet_column] = df[source_sheet_column].replace('nan', 'Unknown')
    
    # Clean notes column
    if notes_column in df.columns:
        df[notes_column] = df[notes_column].astype(str).str.strip()
        df[notes_column] = df[notes_column].replace('nan', '')
    
    # Create mapping dictionary
    categorization_mapping = {}
    duplicate_descriptions = []
    
    for index, row in df.iterrows():
        description = str(row[description_column]).lower()  # Normalize to lowercase
        category = str(row[category_column])
        
        # Check for duplicates
        if description in categorization_mapping:
            duplicate_descriptions.append(description)
            logger.warning(f"Duplicate description found: '{description}' with categories '{categorization_mapping[description]}' and '{category}'")
            # Keep the last occurrence (most recent)
        
        categorization_mapping[description] = category
    
    # Log summary
    logger.info(f"Successfully processed {len(categorization_mapping)} manual categorizations")
    if duplicate_descriptions:
        logger.warning(f"Found {len(set(duplicate_descriptions))} duplicate descriptions")
    
    # Validate mapping quality
    _validate_categorization_mapping(categorization_mapping)
    
    return categorization_mapping


def _validate_categorization_mapping(mapping: Dict[str, str]) -> None:
    """
    Validate the categorization mapping for quality issues
    
    Args:
        mapping: Dictionary mapping descriptions to categories
        
    Raises:
        ValueError: If validation fails
    """
    if not mapping:
        logger.warning("Empty categorization mapping")
        return
    
    # Check for empty or invalid categories
    invalid_categories = []
    for description, category in mapping.items():
        if not category or category.strip() == '':
            invalid_categories.append(description)
    
    if invalid_categories:
        logger.warning(f"Found {len(invalid_categories)} descriptions with empty categories")
        for desc in invalid_categories[:5]:  # Show first 5
            logger.warning(f"  Empty category for: '{desc}'")
    
    # Check for very short descriptions (potential data quality issues)
    short_descriptions = [desc for desc in mapping.keys() if len(desc.strip()) < 3]
    if short_descriptions:
        logger.warning(f"Found {len(short_descriptions)} very short descriptions (< 3 chars)")
        for desc in short_descriptions[:5]:  # Show first 5
            logger.warning(f"  Very short description: '{desc}'")
    
    # Check category distribution
    category_counts = {}
    for category in mapping.values():
        category_counts[category] = category_counts.get(category, 0) + 1
    
    logger.info(f"Category distribution: {category_counts}")
    
    # Warn if any category has very few items (potential typos)
    for category, count in category_counts.items():
        if count == 1:
            logger.info(f"Category '{category}' has only 1 item - verify spelling")


def validate_excel_file_structure(filepath: Path) -> Dict[str, Any]:
    """
    Validate the structure of a manual categorization Excel file
    
    Args:
        filepath: Path to the Excel file
        
    Returns:
        Dictionary with validation results
    """
    logger.info(f"Validating Excel file structure: {filepath}")
    
    validation_result = {
        'is_valid': False,
        'errors': [],
        'warnings': [],
        'file_info': {},
        'sheet_info': {}
    }
    
    try:
        # Check if file exists
        if not filepath.exists():
            validation_result['errors'].append(f"File not found: {filepath}")
            return validation_result
        
        # Load workbook
        wb = openpyxl.load_workbook(filepath, read_only=True)
        
        # Check required sheets
        required_sheets = ["Categorization", "Instructions"]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in wb.sheetnames]
        
        if missing_sheets:
            validation_result['errors'].append(f"Missing required sheets: {missing_sheets}")
            validation_result['errors'].append(f"Available sheets: {wb.sheetnames}")
        else:
            validation_result['sheet_info']['available_sheets'] = wb.sheetnames
        
        # Check Categorization sheet structure
        if "Categorization" in wb.sheetnames:
            ws = wb["Categorization"]
            
            # Check headers
            expected_headers = ["Description", "Source_Sheet", "Frequency", "Category", "Notes"]
            actual_headers = []
            
            for col in range(1, 6):  # First 5 columns
                cell_value = ws.cell(row=1, column=col).value
                actual_headers.append(str(cell_value) if cell_value else "")
            
            missing_headers = [h for h in expected_headers if h not in actual_headers]
            if missing_headers:
                validation_result['errors'].append(f"Missing headers: {missing_headers}")
                validation_result['warnings'].append(f"Actual headers: {actual_headers}")
            else:
                validation_result['sheet_info']['headers'] = actual_headers
            
            # Check data rows
            data_row_count = 0
            for row in range(2, ws.max_row + 1):
                description = ws.cell(row=row, column=1).value
                if description and str(description).strip():
                    data_row_count += 1
            
            validation_result['sheet_info']['data_rows'] = data_row_count
            
            if data_row_count == 0:
                validation_result['warnings'].append("No data rows found in Categorization sheet")
        
        # File information
        validation_result['file_info'] = {
            'file_size': filepath.stat().st_size,
            'last_modified': datetime.fromtimestamp(filepath.stat().st_mtime),
            'file_path': str(filepath)
        }
        
        # Determine if file is valid
        validation_result['is_valid'] = len(validation_result['errors']) == 0
        
        logger.info(f"Validation completed. Valid: {validation_result['is_valid']}")
        if validation_result['errors']:
            logger.error(f"Validation errors: {validation_result['errors']}")
        if validation_result['warnings']:
            logger.warning(f"Validation warnings: {validation_result['warnings']}")
        
        return validation_result
        
    except Exception as e:
        validation_result['errors'].append(f"Error during validation: {e}")
        logger.error(f"Validation error: {e}")
        return validation_result


def get_categorization_statistics(mapping: Dict[str, str]) -> Dict[str, Any]:
    """
    Generate statistics about manual categorizations
    
    Args:
        mapping: Dictionary mapping descriptions to categories
        
    Returns:
        Dictionary with statistics
    """
    if not mapping:
        return {
            'total_categorizations': 0,
            'unique_categories': 0,
            'category_distribution': {},
            'average_description_length': 0,
            'shortest_description': '',
            'longest_description': ''
        }
    
    # Basic counts
    total_categorizations = len(mapping)
    unique_categories = len(set(mapping.values()))
    
    # Category distribution
    category_distribution = {}
    for category in mapping.values():
        category_distribution[category] = category_distribution.get(category, 0) + 1
    
    # Description length statistics
    description_lengths = [len(desc) for desc in mapping.keys()]
    avg_length = sum(description_lengths) / len(description_lengths)
    
    shortest_desc = min(mapping.keys(), key=len)
    longest_desc = max(mapping.keys(), key=len)
    
    return {
        'total_categorizations': total_categorizations,
        'unique_categories': unique_categories,
        'category_distribution': category_distribution,
        'average_description_length': round(avg_length, 2),
        'shortest_description': shortest_desc,
        'longest_description': longest_desc,
        'description_length_range': f"{min(description_lengths)} - {max(description_lengths)}"
    } 