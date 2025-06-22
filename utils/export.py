"""
Excel and Data Export Utilities for BOQ Tools
Handles export of processed BOQ data to various formats like Excel, CSV, and JSON.
"""

import logging
import json
from pathlib import Path
from typing import Dict, List, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

logger = logging.getLogger(__name__)


class ExcelExporter:
    """
    Handles exporting processed BOQ data to different file formats.
    """

    def __init__(self):
        """Initialize the exporter."""
        logger.info("ExcelExporter initialized.")

    def export_data(self, file_mapping: Dict[str, Any], export_path: Path, format_type: str) -> bool:
        """
        Export data to the specified format.

        Args:
            file_mapping: The processed file mapping data.
            export_path: The destination path for the exported file.
            format_type: The format to export to ('normalized_excel', 'summary_excel', 'json', 'csv').

        Returns:
            True if export was successful, False otherwise.
        """
        export_methods = {
            'normalized_excel': self.export_normalized_boq,
            'summary_excel': self.export_summary_report,
            'json': self.export_to_json,
            'csv': self.export_to_csv,
        }

        method = export_methods.get(format_type)
        if not method:
            logger.error(f"Unknown export format: {format_type}")
            return False

        try:
            return method(file_mapping, export_path)
        except Exception as e:
            logger.error(f"Failed to export to {format_type}: {e}", exc_info=True)
            return False

    def export_to_json(self, file_mapping: Dict[str, Any], export_path: Path) -> bool:
        """Export mapping data to a JSON file."""
        export_path.parent.mkdir(parents=True, exist_ok=True)
        with open(export_path, 'w', encoding='utf-8') as f:
            json.dump(file_mapping, f, indent=2, ensure_ascii=False)
        logger.info(f"Successfully exported data to JSON: {export_path}")
        return True

    def export_to_csv(self, file_mapping: Dict[str, Any], export_path: Path) -> bool:
        """Export normalized data to a CSV file."""
        df = self._create_dataframe(file_mapping)
        export_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(export_path, index=False, encoding='utf-8')
        logger.info(f"Successfully exported data to CSV: {export_path}")
        return True

    def export_normalized_boq(self, file_mapping: Dict[str, Any], export_path: Path) -> bool:
        """Export normalized BOQ data to an Excel file."""
        df = self._create_dataframe(file_mapping)
        export_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Normalized BOQ', index=False)
            # Formatting
            ws = writer.sheets['Normalized BOQ']
            self._format_excel_sheet(ws)
        logger.info(f"Successfully exported normalized BOQ to Excel: {export_path}")
        return True

    def export_summary_report(self, file_mapping: Dict[str, Any], export_path: Path) -> bool:
        """Export a summary report to an Excel file."""
        # This can be expanded with more summary details
        summary_data = []
        for sheet_name, sheet_data in file_mapping.get('sheets', {}).items():
            item_count = len(sheet_data.get('items', []))
            total_value = sum(
                float(item.get('total_price', 0) or 0)
                for item in sheet_data.get('items', [])
            )
            summary_data.append({
                "Sheet Name": sheet_name,
                "Item Count": item_count,
                "Total Value": total_value,
            })
        df = pd.DataFrame(summary_data)
        export_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Summary Report', index=False)
            ws = writer.sheets['Summary Report']
            self._format_excel_sheet(ws)
        logger.info(f"Successfully exported summary report to Excel: {export_path}")
        return True
        
    def _create_dataframe(self, file_mapping: Dict[str, Any]) -> pd.DataFrame:
        """Create a pandas DataFrame from the file mapping."""
        records = []
        for sheet_name, sheet_data in file_mapping.get('sheets', {}).items():
            for item in sheet_data.get('items', []):
                record = {
                    "Item No": item.get('item_no'),
                    "Description": item.get('description'),
                    "Unit": item.get('unit'),
                    "Quantity": item.get('quantity'),
                    "Unit Price": item.get('unit_price'),
                    "Total Price": item.get('total_price'),
                    "Category": item.get('category'),
                    "Source Sheet": sheet_name,
                }
                records.append(record)
        return pd.DataFrame(records)

    def _format_excel_sheet(self, ws):
        """Apply standard formatting to an Excel worksheet."""
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill

        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width 