"""
BOQ Processor - Main processing logic for Bill of Quantities
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Any

from .file_processor import ExcelProcessor, SheetMetadata, ContentSample
from utils.config import get_config, ColumnType

logger = logging.getLogger(__name__)


class BOQProcessor:
    """
    Main BOQ processing class that coordinates file processing and analysis
    """
    
    def __init__(self):
        """Initialize the BOQ processor"""
        self.config = get_config()
        self.file_processor: Optional[ExcelProcessor] = None
        self.current_file: Optional[Path] = None
        self.sheets_metadata: Dict[str, SheetMetadata] = {}
        self.sheets_content: Dict[str, ContentSample] = {}
        
        logger.info("BOQ Processor initialized")
    
    def load_excel(self, file_path: Path, sheets_to_process: Optional[List[str]] = None) -> bool:
        """
        Load an Excel file and extract metadata
        
        Args:
            file_path: Path to the Excel file
            sheets_to_process: Optional list of sheet names to process. If None, processes all visible sheets.
            
        Returns:
            True if file loaded successfully, False otherwise
        """
        try:
            # Initialize file processor
            self.file_processor = ExcelProcessor(
                max_memory_mb=self.config.processing_limits.memory_limit_mb
            )
            
            # Load the file
            if not self.file_processor.load_file(file_path):
                logger.error("Failed to load Excel file")
                return False
            
            self.current_file = file_path
            
            # Get visible sheets
            visible_sheets = self.file_processor.get_visible_sheets()
            if not visible_sheets:
                logger.warning("No visible sheets found in file")
                return False
            
            # Filter sheets to process
            if sheets_to_process:
                # Only process specified sheets that are visible
                sheets_to_process = [sheet for sheet in sheets_to_process if sheet in visible_sheets]
                if not sheets_to_process:
                    logger.warning("None of the specified sheets are visible in the file")
                    return False
                logger.info(f"Processing {len(sheets_to_process)} specified sheets out of {len(visible_sheets)} visible sheets")
            else:
                # Process all visible sheets
                sheets_to_process = visible_sheets
                logger.info(f"Processing all {len(sheets_to_process)} visible sheets")
            
            # Extract metadata for specified sheets only
            self.sheets_metadata = {}
            for sheet_name in sheets_to_process:
                try:
                    metadata = self.file_processor.get_sheet_metadata(sheet_name)
                    self.sheets_metadata[sheet_name] = metadata
                except Exception as e:
                    logger.warning(f"Failed to extract metadata for sheet '{sheet_name}': {e}")
            
            # Sample content from specified sheets only
            for sheet_name in sheets_to_process:
                try:
                    sample = self.file_processor.sample_sheet_content(sheet_name, rows=20)
                    self.sheets_content[sheet_name] = sample
                except Exception as e:
                    logger.warning(f"Failed to sample content from sheet '{sheet_name}': {e}")
            
            logger.info(f"Successfully loaded {len(sheets_to_process)} sheets")
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            self._cleanup()
            return False
    
    def process(self) -> Dict[str, Any]:
        """
        Process the loaded Excel file and extract BOQ information
        
        Returns:
            Dictionary containing processed BOQ data
        """
        if not self.file_processor or not self.current_file:
            raise RuntimeError("No file loaded. Call load_excel() first.")
        
        logger.info("Starting BOQ processing")
        
        results = {
            "file_path": str(self.current_file),
            "file_format": self.file_processor.file_format,
            "sheets_processed": 0,
            "sheets_data": {},
            "column_mappings": {},
            "validation_results": {},
            "summary": {}
        }
        
        # Process only the sheets that were loaded
        sheets_to_process = list(self.sheets_metadata.keys())
        
        for sheet_name in sheets_to_process:
            try:
                sheet_data = self._process_sheet(sheet_name)
                if sheet_data:
                    results["sheets_data"][sheet_name] = sheet_data
                    results["sheets_processed"] += 1
                    
            except Exception as e:
                logger.error(f"Error processing sheet '{sheet_name}': {e}")
                results["validation_results"][sheet_name] = {
                    "error": str(e),
                    "status": "failed"
                }
        
        # Generate summary
        results["summary"] = self._generate_summary(results)
        
        logger.info(f"BOQ processing completed. Processed {results['sheets_processed']} sheets")
        return results
    
    def _process_sheet(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """
        Process a single sheet and extract BOQ data
        
        Args:
            sheet_name: Name of the sheet to process
            
        Returns:
            Dictionary with processed sheet data or None if processing failed
        """
        logger.debug(f"Processing sheet: {sheet_name}")
        
        # Get sheet metadata and content
        metadata = self.sheets_metadata.get(sheet_name)
        content = self.sheets_content.get(sheet_name)
        
        if not metadata or not content:
            logger.warning(f"Missing metadata or content for sheet '{sheet_name}'")
            return None
        
        # Classify sheet type
        sheet_type, confidence = self.config.get_sheet_classification(
            sheet_name, content.headers + [str(cell) for row in content.rows for cell in row]
        )
        
        # Map columns
        column_mappings = self._map_columns(content.headers)
        
        # Validate sheet
        validation = self._validate_sheet(metadata, content, column_mappings)
        
        # Extract BOQ data if validation passes
        boq_data = None
        if validation["is_valid"]:
            boq_data = self._extract_boq_data(content, column_mappings)
        
        return {
            "sheet_type": sheet_type,
            "confidence": confidence,
            "metadata": {
                "row_count": metadata.row_count,
                "column_count": metadata.column_count,
                "data_density": metadata.data_density,
                "empty_rows": metadata.empty_rows_count,
                "empty_columns": metadata.empty_columns_count
            },
            "column_mappings": column_mappings,
            "validation": validation,
            "boq_data": boq_data,
            "sample_content": {
                "headers": content.headers,
                "sample_rows": content.rows[:5]  # First 5 rows only
            }
        }
    
    def _map_columns(self, headers: List[str]) -> Dict[str, str]:
        """
        Map column headers to BOQ column types
        
        Args:
            headers: List of column headers
            
        Returns:
            Dictionary mapping column index to column type
        """
        mappings = {}
        
        for col_idx, header in enumerate(headers):
            if not header or header.strip() == "":
                continue
            
            header_lower = header.lower().strip()
            best_match = None
            best_score = 0
            
            # Check each column type
            for col_type in self.config.get_all_column_types():
                mapping = self.config.get_column_mapping(col_type)
                if mapping:
                    for keyword in mapping.keywords:
                        if keyword.lower() in header_lower:
                            score = mapping.weight
                            if score > best_score:
                                best_score = score
                                best_match = col_type.value
            
            if best_match and best_score >= self.config.validation_thresholds.min_column_confidence:
                mappings[str(col_idx)] = best_match
        
        return mappings
    
    def _validate_sheet(self, metadata: SheetMetadata, content: ContentSample, 
                       column_mappings: Dict[str, str]) -> Dict[str, Any]:
        """
        Validate a sheet against BOQ requirements
        
        Args:
            metadata: Sheet metadata
            content: Sheet content sample
            column_mappings: Column mappings
            
        Returns:
            Validation results dictionary
        """
        validation = {
            "is_valid": True,
            "errors": [],
            "warnings": [],
            "score": 0.0
        }
        
        # Check data density
        if metadata.data_density < 0.1:  # Less than 10% data
            validation["warnings"].append("Low data density")
            validation["score"] -= 0.2
        
        # Check for required columns
        required_columns = self.config.get_required_columns()
        mapped_types = set(column_mappings.values())
        
        for required_col in required_columns:
            if required_col.value not in mapped_types:
                validation["errors"].append(f"Missing required column: {required_col.value}")
                validation["score"] -= 0.3
        
        # Check minimum data rows
        if metadata.last_data_row - metadata.first_data_row < self.config.validation_thresholds.min_data_rows:
            validation["errors"].append("Insufficient data rows")
            validation["score"] -= 0.5
        
        # Check empty rows percentage
        empty_percentage = metadata.empty_rows_count / max(1, metadata.row_count)
        if empty_percentage > self.config.validation_thresholds.max_empty_rows_percentage:
            validation["warnings"].append(f"High percentage of empty rows: {empty_percentage:.1%}")
            validation["score"] -= 0.1
        
        # Calculate final score
        validation["score"] = max(0.0, min(1.0, validation["score"] + 1.0))
        
        # Determine if valid
        validation["is_valid"] = len(validation["errors"]) == 0 and validation["score"] >= 0.5
        
        return validation
    
    def _extract_boq_data(self, content: ContentSample, 
                         column_mappings: Dict[str, str]) -> List[Dict[str, Any]]:
        """
        Extract BOQ data from sheet content
        
        Args:
            content: Sheet content sample
            column_mappings: Column mappings
            
        Returns:
            List of BOQ items
        """
        boq_items = []
        
        # Find column indices for each type
        col_indices = {}
        for col_idx, col_type in column_mappings.items():
            col_indices[col_type] = int(col_idx)
        
        # Process each data row
        for row_idx, row in enumerate(content.rows):
            if not row or all(cell.strip() == "" for cell in row):
                continue
            
            item = {
                "row_index": row_idx + 2,  # +2 because we skip header and start from 0
                "description": "",
                "quantity": None,
                "unit": "",
                "unit_price": None,
                "total_price": None,
                "classification": "",
                "code": "",
                "remarks": ""
            }
            
            # Extract data based on column mappings
            for col_type, col_idx in col_indices.items():
                if col_idx < len(row):
                    value = row[col_idx].strip()
                    
                    if col_type == "description":
                        item["description"] = value
                    elif col_type == "quantity":
                        try:
                            item["quantity"] = float(value) if value else None
                        except ValueError:
                            item["quantity"] = None
                    elif col_type == "unit":
                        item["unit"] = value
                    elif col_type == "unit_price":
                        try:
                            item["unit_price"] = float(value) if value else None
                        except ValueError:
                            item["unit_price"] = None
                    elif col_type == "total_price":
                        try:
                            item["total_price"] = float(value) if value else None
                        except ValueError:
                            item["total_price"] = None
                    elif col_type == "classification":
                        item["classification"] = value
                    elif col_type == "code":
                        item["code"] = value
                    elif col_type == "remarks":
                        item["remarks"] = value
            
            # Only add items with at least a description
            if item["description"]:
                boq_items.append(item)
        
        return boq_items
    
    def _generate_summary(self, results: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate summary statistics from processing results
        
        Args:
            results: Processing results
            
        Returns:
            Summary dictionary
        """
        summary = {
            "total_sheets": len(results["sheets_data"]),
            "valid_sheets": 0,
            "total_items": 0,
            "total_value": 0.0,
            "sheet_types": {},
            "column_coverage": {}
        }
        
        for sheet_name, sheet_data in results["sheets_data"].items():
            if sheet_data["validation"]["is_valid"]:
                summary["valid_sheets"] += 1
            
            # Count sheet types
            sheet_type = sheet_data["sheet_type"]
            summary["sheet_types"][sheet_type] = summary["sheet_types"].get(sheet_type, 0) + 1
            
            # Count column coverage
            for col_type in sheet_data["column_mappings"].values():
                summary["column_coverage"][col_type] = summary["column_coverage"].get(col_type, 0) + 1
            
            # Count BOQ items and total value
            if sheet_data["boq_data"]:
                summary["total_items"] += len(sheet_data["boq_data"])
                for item in sheet_data["boq_data"]:
                    if item["total_price"]:
                        summary["total_value"] += item["total_price"]
        
        return summary
    
    def close(self) -> None:
        """Close the processor and clean up resources"""
        self._cleanup()
    
    def _cleanup(self) -> None:
        """Clean up resources"""
        if self.file_processor:
            self.file_processor.close()
            self.file_processor = None
        
        self.current_file = None
        self.sheets_metadata.clear()
        self.sheets_content.clear()
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close() 

    def identify_boq_sheets(self, sheet_names: List[str]) -> List[str]:
        """
        Identify which sheets are likely to be BOQ sheets based on naming patterns
        
        Args:
            sheet_names: List of sheet names to analyze
            
        Returns:
            List of sheet names that are likely BOQ sheets
        """
        boq_sheets = []
        
        # Common BOQ sheet name patterns
        boq_patterns = [
            'boq', 'bill', 'quantity', 'schedule', 'item', 'work', 'construction',
            'civil', 'electrical', 'mechanical', 'structural', 'finishes',
            'earthwork', 'concrete', 'steel', 'masonry', 'carpentry',
            'plumbing', 'hvac', 'roofing', 'painting', 'landscaping'
        ]
        
        for sheet_name in sheet_names:
            sheet_lower = sheet_name.lower()
            
            # Check if sheet name contains BOQ-related keywords
            is_boq_sheet = any(pattern in sheet_lower for pattern in boq_patterns)
            
            # Also include sheets that are not clearly non-BOQ
            non_boq_patterns = ['summary', 'index', 'contents', 'toc', 'cover', 'title']
            is_non_boq = any(pattern in sheet_lower for pattern in non_boq_patterns)
            
            if is_boq_sheet or not is_non_boq:
                boq_sheets.append(sheet_name)
        
        logger.info(f"Identified {len(boq_sheets)} BOQ sheets out of {len(sheet_names)} total sheets")
        return boq_sheets 