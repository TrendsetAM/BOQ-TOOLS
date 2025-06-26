"""
Manual Categorizer for BOQ Tools
Generates Excel files for manual categorization of unmatched descriptions
"""

import logging
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import openpyxl
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