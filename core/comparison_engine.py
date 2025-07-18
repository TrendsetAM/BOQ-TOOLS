"""
Comparison Engine for BOQ Tools
Handles merging and adding data from Comparison BoQ to the main Dataset
"""

import logging
import pandas as pd
from typing import Dict, List, Any, Optional, Union
from dataclasses import dataclass
import re

logger = logging.getLogger(__name__)


@dataclass
class MergeResult:
    """Result of a merge operation"""
    success: bool
    rows_updated: int
    errors: List[str]
    offer_columns_created: List[str]


class ComparisonEngine:
    """
    Handles comparison operations between Comparison BoQ and Dataset
    """
    
    def __init__(self):
        """Initialize the comparison engine"""
        logger.info("Comparison Engine initialized")
    
    def MERGE(self, comparison_row_data: List[str], dataset_dataframe: pd.DataFrame,
              offer_name: str, column_mapping: Dict[int, Any], 
              row_index: int) -> MergeResult:
        """
        MERGE Function: Writes the values for the current row in Comparison BoQ to the new offer-specific columns in the Dataset
        
        This function only updates the new columns created to host the new offer values, 
        not the base columns (description, code, unit, etc.).
        
        Args:
            comparison_row_data: Row data from Comparison BoQ
            dataset_dataframe: Pandas DataFrame representing the Dataset
            offer_name: Name of the offer (used to create column names)
            column_mapping: Dictionary mapping column index to ColumnType
            row_index: Index of the row in the Dataset to update
            
        Returns:
            MergeResult with operation details
        """
        try:
            errors = []
            rows_updated = 0
            offer_columns_created = []
            
            # Validate inputs
            if not comparison_row_data:
                errors.append("No comparison row data provided")
                return MergeResult(False, 0, errors, [])
            
            if dataset_dataframe.empty:
                errors.append("Dataset DataFrame is empty")
                return MergeResult(False, 0, errors, [])
            
            if row_index >= len(dataset_dataframe):
                errors.append(f"Row index {row_index} is out of bounds for DataFrame with {len(dataset_dataframe)} rows")
                return MergeResult(False, 0, errors, [])
            
            # Define the offer-specific columns to create/update
            offer_columns = {
                'quantity': f'quantity[{offer_name}]',
                'unit_price': f'unit_price[{offer_name}]',
                'total_price': f'total_price[{offer_name}]',
                'manhours': f'manhours[{offer_name}]',
                'wage': f'wage[{offer_name}]'
            }
            
            # Create offer-specific columns if they don't exist
            for col_name, offer_col_name in offer_columns.items():
                if offer_col_name not in dataset_dataframe.columns:
                    dataset_dataframe[offer_col_name] = None
                    offer_columns_created.append(offer_col_name)
                    logger.info(f"Created new column: {offer_col_name}")
            
            # Map comparison row data to offer columns
            updated_values = {}
            
            for col_idx, col_type in column_mapping.items():
                if col_idx < len(comparison_row_data):
                    cell_value = comparison_row_data[col_idx].strip() if comparison_row_data[col_idx] else ""
                    
                    # Map to appropriate offer column based on column type
                    if col_type == "QUANTITY" and cell_value:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            if numeric_value is not None:
                                updated_values[offer_columns['quantity']] = numeric_value
                        except Exception as e:
                            errors.append(f"Error converting quantity '{cell_value}': {e}")
                    
                    elif col_type == "UNIT_PRICE" and cell_value:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            if numeric_value is not None:
                                updated_values[offer_columns['unit_price']] = numeric_value
                        except Exception as e:
                            errors.append(f"Error converting unit price '{cell_value}': {e}")
                    
                    elif col_type == "TOTAL_PRICE" and cell_value:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            if numeric_value is not None:
                                updated_values[offer_columns['total_price']] = numeric_value
                        except Exception as e:
                            errors.append(f"Error converting total price '{cell_value}': {e}")
                    
                    elif col_type == "MANHOURS" and cell_value:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            if numeric_value is not None:
                                updated_values[offer_columns['manhours']] = numeric_value
                        except Exception as e:
                            errors.append(f"Error converting manhours '{cell_value}': {e}")
                    
                    elif col_type == "WAGE" and cell_value:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            if numeric_value is not None:
                                updated_values[offer_columns['wage']] = numeric_value
                        except Exception as e:
                            errors.append(f"Error converting wage '{cell_value}': {e}")
            
            # Update the DataFrame with the new values
            if updated_values:
                for col_name, value in updated_values.items():
                    dataset_dataframe.at[row_index, col_name] = value
                
                rows_updated = 1
                logger.info(f"Updated row {row_index} with {len(updated_values)} offer-specific values")
            else:
                logger.warning(f"No valid values found to merge for row {row_index}")
            
            return MergeResult(
                success=len(errors) == 0,
                rows_updated=rows_updated,
                errors=errors,
                offer_columns_created=offer_columns_created
            )
            
        except Exception as e:
            error_msg = f"Error in MERGE function: {e}"
            logger.error(error_msg)
            return MergeResult(False, 0, [error_msg], [])
    
    def ADD(self, comparison_row_data: List[str], dataset_dataframe: pd.DataFrame,
            column_mapping: Dict[int, Any], position: int) -> Dict[str, Any]:
        """
        ADD Function: Appends the current Comparison BoQ row to the Dataset, providing all necessary values for all columns in the Dataset
        
        Only Categories should be skipped as the row has not yet been categorized.
        
        Args:
            comparison_row_data: Row data from Comparison BoQ
            dataset_dataframe: Pandas DataFrame representing the Dataset
            column_mapping: Dictionary mapping column index to ColumnType
            position: Position number for the new row
            
        Returns:
            Dictionary with operation details including success status and any errors
        """
        try:
            errors = []
            new_row_data = {}
            
            # Validate inputs
            if not comparison_row_data:
                errors.append("No comparison row data provided")
                return {"success": False, "errors": errors, "row_added": False}
            
            if dataset_dataframe is None:
                errors.append("Dataset DataFrame is None")
                return {"success": False, "errors": errors, "row_added": False}
            
            # Map comparison row data to dataset columns
            for col_idx, col_type in column_mapping.items():
                if col_idx < len(comparison_row_data):
                    cell_value = comparison_row_data[col_idx].strip() if comparison_row_data[col_idx] else ""
                    
                    # Skip category assignment (to be handled by RECATEGORIZATION)
                    if col_type == "CATEGORY":
                        new_row_data[col_type] = ""  # Empty category
                        continue
                    
                    # Handle different column types
                    if col_type in ["QUANTITY", "UNIT_PRICE", "TOTAL_PRICE", "MANHOURS", "WAGE"]:
                        try:
                            numeric_value = self._convert_to_numeric(cell_value)
                            new_row_data[col_type] = numeric_value if numeric_value is not None else ""
                        except Exception as e:
                            errors.append(f"Error converting {col_type} '{cell_value}': {e}")
                            new_row_data[col_type] = ""
                    
                    elif col_type in ["DESCRIPTION", "CODE", "UNIT", "NOTES"]:
                        # Text fields - use as is
                        new_row_data[col_type] = cell_value
                    
                    else:
                        # Unknown column type - use as is
                        new_row_data[col_type] = cell_value
            
            # Add position to the new row data
            new_row_data["Position"] = position
            
            # Create a new row as a dictionary
            if new_row_data:
                # Map the data to the actual DataFrame columns
                row_data_for_df = {}
                for col in dataset_dataframe.columns:
                    if col in new_row_data:
                        row_data_for_df[col] = new_row_data[col]
                    else:
                        row_data_for_df[col] = ""  # Fill missing columns with empty string
                
                # Append to DataFrame in place
                dataset_dataframe.loc[len(dataset_dataframe)] = row_data_for_df
                
                logger.info(f"Added new row with position {position} to dataset")
                
                return {
                    "success": len(errors) == 0,
                    "errors": errors,
                    "row_added": True,
                    "new_row_index": len(dataset_dataframe) - 1,
                    "position": position
                }
            else:
                errors.append("No valid data found to add")
                return {"success": False, "errors": errors, "row_added": False}
                
        except Exception as e:
            error_msg = f"Error in ADD function: {e}"
            logger.error(error_msg)
            return {"success": False, "errors": [error_msg], "row_added": False}
    
    def _convert_to_numeric(self, value: str) -> Optional[Union[int, float]]:
        """
        Convert string value to numeric, handling various formats
        
        Args:
            value: String value to convert
            
        Returns:
            Numeric value (int or float) or None if conversion fails
        """
        try:
            if not value or value.strip() == "":
                return None
            
            # Remove currency symbols, commas, and whitespace
            clean_value = re.sub(r'[\$€£¥₹,\s\u00A0]', '', value.strip())
            
            # Handle European decimal format (comma as decimal separator)
            if ',' in clean_value and clean_value.count(',') == 1 and '.' not in clean_value:
                clean_value = clean_value.replace(',', '.')
            
            # Convert to float first
            numeric_value = float(clean_value)
            
            # Return as int if it's a whole number, otherwise as float
            if numeric_value.is_integer():
                return int(numeric_value)
            else:
                return numeric_value
                
        except (ValueError, TypeError) as e:
            logger.warning(f"Failed to convert '{value}' to numeric: {e}")
            return None
    
    def validate_merge_operation(self, dataset_dataframe: pd.DataFrame, 
                               offer_name: str, row_index: int) -> bool:
        """
        Validate that a merge operation can be performed
        
        Args:
            dataset_dataframe: Dataset DataFrame
            offer_name: Name of the offer
            row_index: Index of the row to update
            
        Returns:
            True if merge operation is valid, False otherwise
        """
        try:
            # Check if DataFrame exists and is not empty
            if dataset_dataframe is None or dataset_dataframe.empty:
                logger.error("Dataset DataFrame is None or empty")
                return False
            
            # Check if row index is valid
            if row_index < 0 or row_index >= len(dataset_dataframe):
                logger.error(f"Row index {row_index} is out of bounds")
                return False
            
            # Check if offer name is valid
            if not offer_name or not offer_name.strip():
                logger.error("Offer name is empty or invalid")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Error validating merge operation: {e}")
            return False
    
    def get_offer_columns(self, offer_name: str) -> Dict[str, str]:
        """
        Get the offer-specific column names for a given offer
        
        Args:
            offer_name: Name of the offer
            
        Returns:
            Dictionary mapping base column names to offer-specific column names
        """
        return {
            'quantity': f'quantity[{offer_name}]',
            'unit_price': f'unit_price[{offer_name}]',
            'total_price': f'total_price[{offer_name}]',
            'manhours': f'manhours[{offer_name}]',
            'wage': f'wage[{offer_name}]'
        } 