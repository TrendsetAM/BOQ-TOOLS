"""
Comparison Engine for BOQ Tools
Handles merging and adding data from Comparison BoQ to the main Dataset
"""

import logging
import pandas as pd
from typing import Dict, List, Any, Optional, Union, Tuple # Added Tuple
from dataclasses import dataclass
import re
from core.validator import ValidationIssue, ValidationLevel, ValidationType

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
        # logger.info("Comparison Engine initialized")
    
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
            from utils.config import ColumnType
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
            
            # CRITICAL FIX: Always create offer-specific columns and transfer data
            # Remove conditional logic that was preventing data transfer for empty/zero values
            for col_name, offer_col_name in offer_columns.items():
                if offer_col_name not in dataset_dataframe.columns:
                    dataset_dataframe[offer_col_name] = 0  # Initialize with 0 instead of None
                    offer_columns_created.append(offer_col_name)
                    logger.info(f"Created new offer column: {offer_col_name}")
            
            # CRITICAL FIX: Always attempt to map and transfer data, regardless of empty values
            # Map comparison row data to offer columns
            updated_values = {}
            
            for col_idx, col_type in column_mapping.items():
                if col_idx < len(comparison_row_data):
                    cell_value = str(comparison_row_data[col_idx]).strip() if comparison_row_data[col_idx] is not None else ""
                    
                    # CRITICAL FIX: Remove the "and cell_value" condition - always attempt conversion
                    # Map to appropriate offer column based on column type
                    if col_type == "QUANTITY":
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, row_index, col_type, offer_name # Using offer_name as source_sheet for MERGE context
                        )
                        if not is_valid:
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=row_index,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{offer_name}' at row {row_index+2}"
                                )
                            )
                        updated_values[offer_columns['quantity']] = numeric_value
                    
                    elif col_type == "UNIT_PRICE":
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, row_index, col_type, offer_name
                        )
                        if not is_valid:
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=row_index,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{offer_name}' at row {row_index+2}"
                                )
                            )
                        updated_values[offer_columns['unit_price']] = numeric_value
                    
                    elif col_type == "TOTAL_PRICE":
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, row_index, col_type, offer_name
                        )
                        if not is_valid:
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=row_index,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{offer_name}' at row {row_index+2}"
                                )
                            )
                        updated_values[offer_columns['total_price']] = numeric_value
                    
                    elif col_type == "MANHOURS":
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, row_index, col_type, offer_name
                        )
                        if not is_valid:
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=row_index,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{offer_name}' at row {row_index+2}"
                                )
                            )
                        updated_values[offer_columns['manhours']] = float(numeric_value) if numeric_value != 0 else 0.0
                    
                    elif col_type == "WAGE":
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, row_index, col_type, offer_name
                        )
                        if not is_valid:
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=row_index,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{offer_name}' at row {row_index+2}"
                                )
                            )
                        updated_values[offer_columns['wage']] = numeric_value
            
            # CRITICAL FIX: Always update the DataFrame, even with zero values
            # Update the DataFrame with the new values
            for col_name, value in updated_values.items():
                dataset_dataframe.at[row_index, col_name] = value
            
            rows_updated = 1 if updated_values else 0
            logger.info(f"MERGE: Updated row {row_index} with {len(updated_values)} offer-specific values")
            
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
            column_mapping: Dict[int, Any], position: int, offer_name: str = None, source_sheet: str = None) -> Dict[str, Any]:
        """
        ADD Function: Appends the current Comparison BoQ row to the Dataset, providing all necessary values for all columns in the Dataset
        
        Populates both master columns (Description, Code, Unit, Scope, Source_Sheet) and offer-specific columns (Quantity[offer], Unit Price[offer], etc.)
        
        Args:
            comparison_row_data: Row data from Comparison BoQ
            dataset_dataframe: Pandas DataFrame representing the Dataset
            column_mapping: Dictionary mapping column index to ColumnType
            position: Position number for the new row
            offer_name: Name of the offer for creating offer-specific columns
            source_sheet: Name of the source sheet from comparison BoQ Excel file
            
        Returns:
            Dictionary with operation details including success status and any errors
        """
        try:
            logger.error("=== ENTERING ADD FUNCTION ===")
            logger.error(f"comparison_row_data: {comparison_row_data}")
            logger.error(f"column_mapping: {column_mapping}")
            logger.error(f"position: {position}")
            logger.error(f"offer_name: {offer_name}")
            logger.error(f"source_sheet: {source_sheet}")
        except Exception as debug_error:
            # If even logging fails, return an error
            return {"success": False, "errors": [f"Debug logging failed: {debug_error}"], "row_added": False}
        
        try:
            from utils.config import ColumnType
            errors = []
            new_row_data = {}
            
            # Validate inputs
            if not comparison_row_data:
                errors.append("No comparison row data provided")
                return {"success": False, "errors": errors, "row_added": False}
            
            if dataset_dataframe is None:
                errors.append("Dataset DataFrame is None")
                return {"success": False, "errors": errors, "row_added": False}
            
            # Initialize new_row_data with default values for all columns in dataset_dataframe
            for col in dataset_dataframe.columns:
                # Check if it's a numeric master column (not an offer-specific column)
                if (col in ['quantity', 'unit_price', 'total_price', 'manhours', 'wage'] and '[' not in col):
                    new_row_data[col] = 0  # Default to 0 for numeric master columns
                elif col == "Position":
                    new_row_data[col] = position # Set position directly
                elif col == "Source_Sheet" and source_sheet:
                    new_row_data[col] = source_sheet  # Set source sheet name
                else:
                    new_row_data[col] = ""  # Default to empty string for other columns
            
            # DEBUG: Log the inputs to understand what's happening
            logger.warning(f"ADD DEBUG: comparison_row_data = {comparison_row_data}")
            logger.warning(f"ADD DEBUG: column_mapping = {column_mapping}")
            logger.warning(f"ADD DEBUG: dataset_dataframe.columns = {list(dataset_dataframe.columns)}")
            
            # Define offer-specific columns if offer_name is provided
            offer_columns = {}
            if offer_name:
                offer_columns = {
                    'quantity': f'quantity[{offer_name}]',
                    'unit_price': f'unit_price[{offer_name}]',
                    'total_price': f'total_price[{offer_name}]',
                    'manhours': f'manhours[{offer_name}]',
                    'wage': f'wage[{offer_name}]'
                }
                logger.warning(f"ADD DEBUG: offer_columns = {offer_columns}")
            
            # Map ColumnType strings to actual DataFrame column names
            column_type_to_df_column = {
                "DESCRIPTION": "Description",
                "CODE": "code",
                "UNIT": "unit",
                "QUANTITY": "quantity",
                "UNIT_PRICE": "unit_price",
                "TOTAL_PRICE": "total_price",
                "MANHOURS": "manhours",
                "WAGE": "wage",
                "CATEGORY": "Category",
                "SCOPE": "scope"
            }
            
            # Process each column in the comparison data
            for col_idx, col_type in column_mapping.items():
                if col_idx < len(comparison_row_data):
                    cell_value = str(comparison_row_data[col_idx]).strip() if comparison_row_data[col_idx] is not None else ""
                    
                    # Skip unmapped columns (those without a proper column type)
                    if not col_type or col_type not in column_type_to_df_column:
                        logger.warning(f"ADD DEBUG: Skipped col_idx={col_idx}, col_type='{col_type}', cell_value='{cell_value}' (unmapped column)")
                        continue
                    
                    # Get the actual DataFrame column name
                    df_column_name = column_type_to_df_column.get(col_type, col_type)
                    
                    logger.warning(f"ADD DEBUG: col_idx={col_idx}, col_type='{col_type}', cell_value='{cell_value}', df_column_name='{df_column_name}'")
                    
                    # Skip category assignment (to be handled by RECATEGORIZATION)
                    if col_type == "CATEGORY":
                        if df_column_name in new_row_data:
                            new_row_data[df_column_name] = ""  # Empty category
                            logger.warning(f"ADD DEBUG: Set {df_column_name} = '' (category)")
                        continue
                    
                    # Handle numeric columns - put values in OFFER-SPECIFIC columns, keep master columns as 0
                    if col_type in ["QUANTITY", "UNIT_PRICE", "TOTAL_PRICE", "MANHOURS", "WAGE"]:
                        numeric_value, is_valid, error_msg = self._convert_to_numeric(
                            cell_value, position, col_type, source_sheet # Using position as row_index for ADD context
                        )
                        if not is_valid:
                            errors.append(error_msg) # Add to ADD function's errors
                            self.comparison_warnings.append(
                                ValidationIssue(
                                    row_index=position,
                                    column_index=col_idx,
                                    validation_type=ValidationType.DATA_TYPE,
                                    level=ValidationLevel.WARNING,
                                    message=error_msg,
                                    expected_value="Numeric",
                                    actual_value=cell_value,
                                    suggestion=f"Correct format for '{col_type}' in sheet '{source_sheet}' at row {position+2}"
                                )
                            )
                        
                        # Put numeric values in offer-specific columns if available
                        if offer_name and col_type.lower() in ['quantity', 'unit_price', 'total_price', 'manhours', 'wage']:
                            offer_col_name = offer_columns.get(col_type.lower())
                            if offer_col_name and offer_col_name in new_row_data:
                                new_row_data[offer_col_name] = numeric_value
                                logger.warning(f"ADD DEBUG: Set {offer_col_name} = {numeric_value} (offer-specific)")
                        
                        # Keep master numeric columns as 0 (they were initialized above)
                        logger.warning(f"ADD DEBUG: Master {df_column_name} remains 0 (numeric master)")
                    
                    # Handle text columns - put values in master columns
                    elif col_type in ["DESCRIPTION", "CODE", "UNIT", "SCOPE"]:
                        if df_column_name in new_row_data and cell_value:  # Only set if not empty
                            # Special handling for DESCRIPTION - use the longest/most descriptive text
                            if col_type == "DESCRIPTION":
                                # For ADD rows, always use the description from the comparison data
                                new_row_data[df_column_name] = cell_value
                                logger.warning(f"ADD DEBUG: Set {df_column_name} = '{cell_value}' (from comparison data)")
                            # Special handling for SCOPE - ignore "nan" values
                            elif col_type == "SCOPE":
                                if cell_value.lower() != "nan":
                                    new_row_data[df_column_name] = cell_value
                                    logger.warning(f"ADD DEBUG: Set {df_column_name} = '{cell_value}' (scope)")
                                else:
                                    logger.warning(f"ADD DEBUG: Skipped {df_column_name} = '{cell_value}' (nan value)")
                            # Handle other text columns normally
                            else:
                                new_row_data[df_column_name] = cell_value
                                logger.warning(f"ADD DEBUG: Set {df_column_name} = '{cell_value}' (text)")
                    
                    else:
                        # Unknown column type - skip it to avoid polluting master columns
                        logger.warning(f"ADD DEBUG: Skipped col_idx={col_idx}, col_type='{col_type}', cell_value='{cell_value}' (unknown column type)")
            
            logger.warning(f"ADD DEBUG: Final new_row_data = {new_row_data}")
            
            # Create a new row as a dictionary
            if new_row_data:
                # Append to DataFrame in place
                dataset_dataframe.loc[len(dataset_dataframe)] = new_row_data
                
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
    
    def _convert_to_numeric(self, value: str, row_index: int, column_name: str, source_sheet: str) -> Tuple[float, bool, Optional[str]]:
        """
        Convert string value to numeric, handling various formats.
        Returns a tuple: (numeric_value, success_flag, error_message_or_None)
        """
        try:
            if not value or str(value).strip() == "":
                return 0.0, True, None  # Empty value is valid, treated as 0.0
            
            value_str = str(value).strip()
            
            # Remove currency symbols, commas, and whitespace
            clean_value = re.sub(r'[\$€£¥₹,\s\u00A0]', '', value_str)
            
            # Handle European decimal format (comma as decimal separator)
            if ',' in clean_value and clean_value.count(',') == 1 and '.' not in clean_value:
                clean_value = clean_value.replace(',', '.')
            
            # Handle multiple dots (thousands separators) - keep only the last one as decimal
            if clean_value.count('.') > 1:
                parts = clean_value.split('.')
                clean_value = ''.join(parts[:-1]) + '.' + parts[-1]
            
            numeric_value = float(clean_value)
            return round(numeric_value, 6), True, None
                
        except (ValueError, TypeError) as e:
            error_msg = f"Invalid numeric format for '{column_name}' at row {row_index+2} (Sheet: '{source_sheet}'): '{value}' - {e}"
            logger.warning(error_msg)
            return 0.0, False, error_msg # Return 0.0, False, and error message
    
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

class ComparisonProcessor:
    """
    Orchestrates the entire comparison workflow between a master (reference) BoQ and a comparison BoQ.
    Manages state, coordinates row validation, instance matching, merging, adding, and cleanup.
    """
    def __init__(self):
        # Master/reference BoQ DataFrame
        self.master_dataset = None
        # Comparison BoQ DataFrame
        self.comparison_data = None
        # Set or tracker for manually invalidated rows (row keys)
        self.manual_invalidations = set()
        # Row review/validation results (list of dicts or DataFrame)
        self.row_results = []
        # Optionally, store instance match results, merge/add logs, etc.
        self.instance_matches = []
        self.merge_results = []
        self.add_results = []
        self.comparison_warnings = [] # New: To collect warnings during comparison

    def load_master_dataset(self, df, manual_invalidations=None):
        """
        Load the master/reference BoQ DataFrame and manual invalidations.
        Args:
            df: pandas DataFrame for the master BoQ
            manual_invalidations: set of row keys or tracker (optional)
        """
        self.master_dataset = df
        if manual_invalidations is not None:
            self.manual_invalidations = set(manual_invalidations)
        else:
            self.manual_invalidations = set()
        
        # logger.info(f"Master dataset loaded with {len(df)} rows")
        # logger.info(f"Master dataset columns: {list(df.columns)}")
        # logger.info(f"Master dataset shape: {df.shape}")
        
        # Log some sample descriptions for debugging
        if 'Description' in df.columns:
            sample_descriptions = df['Description'].head(5).tolist()
            # logger.info(f"Master dataset sample descriptions: {sample_descriptions}")
            
            # Check for any descriptions that might be problematic
            if len(df) > 190:  # If we have enough rows to check row 195
                row_195_desc = df.iloc[194]['Description'] if 194 < len(df) else "N/A"
                # logger.info(f"Master dataset row 195 description: '{row_195_desc}'")

    def load_comparison_data(self, df):
        """
        Load the comparison BoQ DataFrame.
        Args:
            df: pandas DataFrame for the comparison BoQ
        """
        self.comparison_data = df
        
        # logger.info(f"Comparison dataset loaded with {len(df)} rows")
        # logger.info(f"Comparison dataset columns: {list(df.columns)}")
        # logger.info(f"Comparison dataset shape: {df.shape}")
        
        # Log some sample descriptions for debugging
        if 'Description' in df.columns:
            sample_descriptions = df['Description'].head(5).tolist()
            # logger.info(f"Comparison dataset sample descriptions: {sample_descriptions}")
            
            # Check for any descriptions that might be problematic
            if len(df) > 190:  # If we have enough rows to check row 195
                row_195_desc = df.iloc[194]['Description'] if 194 < len(df) else "N/A"
                # logger.info(f"Comparison dataset row 195 description: '{row_195_desc}'")

    def validate_comparison_data(self):
        """
        Validate that the comparison data and master dataset are compatible (e.g., columns match).
        Returns:
            bool: True if valid, False otherwise
            str: Error message if not valid
        """
        if self.master_dataset is None or self.comparison_data is None:
            return False, "Master or comparison dataset not loaded."
        
        # Check if master dataset has descriptions
        if 'Description' in self.master_dataset.columns:
            master_descriptions = self.master_dataset['Description'].head(10).tolist()
            empty_count = sum(1 for desc in master_descriptions if not desc or str(desc).strip() == '')
            if empty_count == len(master_descriptions):
                return False, "Master dataset has empty descriptions. Cannot perform comparison without descriptions."
            elif empty_count > 0:
                logger.warning(f"Master dataset has {empty_count}/{len(master_descriptions)} empty descriptions")
            else:
                # logger.info(f"Master dataset has valid descriptions: {len(master_descriptions) - empty_count}/{len(master_descriptions)} non-empty")
                pass
        
        # Check column compatibility
        master_cols = set(self.master_dataset.columns)
        comp_cols = set(self.comparison_data.columns)
        missing_in_comp = master_cols - comp_cols
        if missing_in_comp:
            return False, f"Comparison data missing columns: {missing_in_comp}"
        
        # Optionally, check for required columns
        return True, "Validation successful." 

    def process_comparison_rows(self, row_classifier=None, key_columns=None):
        """
        For each row in the comparison data:
        - If the row matches a manual invalidation (by key), mark as invalid (MANUAL_OVERRIDE)
        - Otherwise, call ROW_VALIDITY
        - Store results for row review
        Args:
            row_classifier: Optional, RowClassifier instance (if not provided, will import and create)
            key_columns: List of columns to use as row key (default: ['Description'])
        Returns:
            List of dicts with row index, key, validity, and reason
        """
        if self.comparison_data is None:
            raise ValueError("Comparison data not loaded.")
        if key_columns is None:
            key_columns = ['Description']
        if row_classifier is None:
            from core.row_classifier import RowClassifier, RowType # Added RowType
            row_classifier = RowClassifier()
        results = []
        for idx, row in self.comparison_data.iterrows():
            # Build row key (tuple of key column values)
            key = tuple(str(row.get(col, '')).strip() for col in key_columns)
            key_str = '|'.join(key)
            description = key[0] if key else ''
            
            # Check manual invalidation - check if description is in any invalidation
            is_manually_invalid = any(description in invalidation.split('|')[0] for invalidation in self.manual_invalidations)
            
            if is_manually_invalid:
                is_valid = False
                reason = 'MANUAL_OVERRIDE'
            else:
                # Prepare column mapping for ROW_VALIDITY
                # Convert string column names to ColumnType enum values
                from utils.config import ColumnType
                column_mapping = {}
                for i, col_name in enumerate(self.comparison_data.columns):
                    try:
                        # Map common column names to ColumnType enum
                        if col_name.lower() in ['description', 'desc', 'item']:
                            column_mapping[i] = ColumnType.DESCRIPTION
                        elif col_name.lower() in ['quantity', 'qty', 'qty.']:
                            column_mapping[i] = ColumnType.QUANTITY
                        elif col_name.lower() in ['unit_price', 'unit price', 'price', 'rate']:
                            column_mapping[i] = ColumnType.UNIT_PRICE
                        elif col_name.lower() in ['total_price', 'total price', 'amount', 'total']:
                            column_mapping[i] = ColumnType.TOTAL_PRICE
                        elif col_name.lower() in ['code', 'item code', 'ref']:
                            column_mapping[i] = ColumnType.CODE
                        elif col_name.lower() in ['unit', 'uom']:
                            column_mapping[i] = ColumnType.UNIT
                        elif col_name.lower() in ['manhours', 'ore/u.m.', 'ore', 'man hours']:
                            column_mapping[i] = ColumnType.MANHOURS
                        elif col_name.lower() in ['wage', 'euro/hour', 'hourly rate']:
                            column_mapping[i] = ColumnType.WAGE
                        else:
                            # Default to DESCRIPTION for unknown columns
                            column_mapping[i] = ColumnType.DESCRIPTION
                    except Exception as e:
                        logger.warning(f"Could not map column {col_name}: {e}")
                        column_mapping[i] = ColumnType.DESCRIPTION
                
                row_values = [str(row[col]) if row[col] is not None else '' for col in self.comparison_data.columns]
                
                # Call classify_rows to get detailed validation results
                # Pass the sheet name for better context in warnings
                row_classification_result = row_classifier.classify_rows(
                    [row_values], # Pass as a list of one row
                    column_mapping,
                    sheet_name=row.get('Source_Sheet', 'Comparison') # Get source sheet from row data
                )
                
                # Extract validity and reason from the classification result
                if row_classification_result.classifications:
                    classification = row_classification_result.classifications[0]
                    # A row is considered valid if it's a PRIMARY_LINE_ITEM and has no ERROR level validation issues
                    is_valid = (classification.row_type == RowType.PRIMARY_LINE_ITEM and
                                not any(issue.level == ValidationLevel.ERROR for issue in classification.validation_errors))
                    reason = classification.row_type.value # Use row_type as reason
                    
                    # Collect all validation issues (errors and warnings)
                    for issue in classification.validation_errors:
                        self.comparison_warnings.append(issue)
                else:
                    is_valid = False
                    reason = 'NO_CLASSIFICATION'
                    self.comparison_warnings.append(
                        ValidationIssue(
                            row_index=idx,
                            column_index=None,
                            validation_type=ValidationType.BUSINESS_RULE,
                            level=ValidationLevel.ERROR,
                            message="Row could not be classified",
                            expected_value="Valid row data",
                            actual_value="No classification",
                            suggestion="Review row content and column mappings"
                        )
                    )
                
            results.append({
                'row_index': idx,
                'key': key_str,
                'is_valid': is_valid,
                'reason': reason
            })
        self.row_results = results
        return results

    def process_valid_rows(self, instance_matcher=None, comparison_engine=None, key_columns=None, offer_name=None):
        """
        For each valid row in the comparison data:
        - Use LIST_INSTANCES to get all instances with the same description in both master and comparison datasets
        - For each nth instance, if both exist, call MERGE; if only in comparison, call ADD
        Args:
            instance_matcher: Optional, InstanceMatcher instance
            comparison_engine: Optional, ComparisonEngine instance
            key_columns: List of columns to use as row key (default: ['Description'])
            offer_name: Name of the offer for creating offer-specific columns
        Returns:
            List of merge/add operation results
        """
        import pandas as pd
        if self.comparison_data is None or self.master_dataset is None:
            raise ValueError("Comparison or master dataset not loaded.")
        if key_columns is None:
            key_columns = ['Description']
        if instance_matcher is None:
            from core.instance_matcher import InstanceMatcher
            instance_matcher = InstanceMatcher()
        if comparison_engine is None:
            comparison_engine = ComparisonEngine()
        
        # CRITICAL FIX 1: Ensure offer-specific columns are always created at the beginning
        # This guarantees columns exist even in 100% ADD scenarios
        if offer_name:
            offer_columns = {
                'quantity': f'quantity[{offer_name}]',
                'unit_price': f'unit_price[{offer_name}]',
                'total_price': f'total_price[{offer_name}]',
                'manhours': f'manhours[{offer_name}]',
                'wage': f'wage[{offer_name}]'
            }
            
            # Create offer columns if they don't exist, initialize with 0
            for col_name, offer_col_name in offer_columns.items():
                if offer_col_name not in self.master_dataset.columns:
                    self.master_dataset[offer_col_name] = 0
                    logger.info(f"Created offer column: {offer_col_name}")
        
        # Filter valid rows from previous step
        valid_rows = [r for r in self.row_results if r['is_valid']]
        merge_results = []
        add_results = []
        
        # Track instance counts per description for debugging
        description_instance_counts = {}
        
        logger.info(f"=== STARTING PROCESS_VALID_ROWS ===")
        logger.info(f"Total valid rows to process: {len(valid_rows)}")
        
        for row_info in valid_rows:
            idx = row_info['row_index']
            row = self.comparison_data.loc[idx]
            
            # Build key for instance matching
            key = tuple(str(row.get(col, '')).strip() for col in key_columns)
            description = key[0] if key else ''
            
            # # Enhanced logging for row 195 specifically
            is_row_195 = (idx == 195)
            if is_row_195:
            #     logger.info(f"=== PROCESSING ROW 195 ===")
            #     logger.info(f"Row 195 description: '{description}'")
            #     logger.info(f"Row 195 key: {key}")
            #     logger.info(f"Row 195 data: {dict(row)}")
                pass
            
            # CRITICAL FIX: Use consistent normalized whitespace matching for both datasets
            # This handles cases where descriptions have different whitespace/newline characters
            normalized_desc = ' '.join(description.split())  # Normalize whitespace
            
            # Apply same normalization to both comparison and master datasets
            comp_instances = self.comparison_data[
                self.comparison_data[key_columns[0]].str.lower().apply(lambda x: ' '.join(str(x).split())) == normalized_desc.lower()
            ]
            master_instances = self.master_dataset[
                self.master_dataset[key_columns[0]].str.lower().apply(lambda x: ' '.join(str(x).split())) == normalized_desc.lower()
            ]
            
            # Debug: Log matching information
            # logger.info(f"Row {idx} - Description: '{description}'")
            # logger.info(f"Row {idx} - Normalized description: '{normalized_desc}'")
            # logger.info(f"Row {idx} - Comparison instances found: {len(comp_instances)}")
            # logger.info(f"Row {idx} - Master instances found: {len(master_instances)}")
            
            if is_row_195:
                # logger.info(f"Row 195 - Comparison instance indices: {list(comp_instances.index)}")
                # logger.info(f"Row 195 - Master instance indices: {list(master_instances.index)}")
                # logger.info(f"Row 195 - Current description instance count: {description_instance_counts.get(description, 0)}")
                pass
            else:
                pass
            
            # If no matches found with normalized matching, try exact matching as fallback
            if len(master_instances) == 0:
                master_instances = self.master_dataset[self.master_dataset[key_columns[0]] == description]
                # logger.info(f"Exact match master instances found: {len(master_instances)}")
                
                # If still no matches, try case-insensitive matching
                if len(master_instances) == 0:
                    master_instances = self.master_dataset[
                        self.master_dataset[key_columns[0]].str.lower() == description.lower()
                    ]
                    # logger.info(f"Case-insensitive master instances found: {len(master_instances)}")
                    
                    # If still no matches, try fuzzy matching for very similar descriptions
                    if len(master_instances) == 0:
                        # Check for descriptions that are very similar (might be encoding issues)
                        master_desc_lower = self.master_dataset[key_columns[0]].str.lower()
                        similar_matches = master_desc_lower[master_desc_lower.str.contains(description.lower()[:50], na=False)]
                        if len(similar_matches) > 0:
                            logger.warning(f"Found {len(similar_matches)} similar descriptions for '{description[:50]}...'")
                            logger.warning(f"Similar descriptions: {similar_matches.head(3).tolist()}")
                        else:
                            pass
                    else:
                        pass
                
                # If still no matches, this is normal - the description doesn't exist in master
                # This means we should ADD this row, not throw an error
                if len(master_instances) == 0:
                    # logger.info(f"No match found for description: '{description}' - will be added as new row")
                    # Continue processing - this row will be added in the ADD section below
                    pass
            # For each nth instance, match and merge/add
            for n, (comp_idx, comp_row) in enumerate(comp_instances.iterrows()):
                # Enhanced logging for instance processing
                if is_row_195:
                    # logger.info(f"Row 195 - Processing instance {n} of {len(comp_instances)}")
                    # logger.info(f"Row 195 - Current instance number: {n}")
                    # logger.info(f"Row 195 - Master instances available: {len(master_instances)}")
                    # logger.info(f"Row 195 - Decision condition: {n} < {len(master_instances)} = {n < len(master_instances)}")
                    pass
                
                if n < len(master_instances):
                    master_idx = master_instances.index[n]
                    # logger.info(f"Row {idx} - MERGE decision: instance {n} -> master row {master_idx}")
                    
                    if is_row_195:
                        # logger.info(f"Row 195 - MERGE: Will merge into master row {master_idx}")
                        pass
                    
                    # Call MERGE: merge comp_row into master_dataset at master_idx
                    # Get units for comparison
                    master_unit = self.master_dataset.loc[master_idx, 'unit']
                    comp_unit = comp_row.get('unit', '') # Get unit from comparison row

                    # Check for unit mismatch (Test 4)
                    if str(master_unit).strip().lower() != str(comp_unit).strip().lower():
                        self.comparison_warnings.append(
                            ValidationIssue(
                                row_index=master_idx, # Master row index
                                column_index=None, # Unit column index is not fixed here
                                validation_type=ValidationType.CONSISTENCY, # Or a new UNIT_MISMATCH type
                                level=ValidationLevel.WARNING,
                                message=f"Unit mismatch for item: Master unit '{master_unit}' vs. Comparison unit '{comp_unit}'",
                                expected_value=master_unit,
                                actual_value=comp_unit,
                                suggestion=f"Review unit for item in sheet '{comp_row.get('Source_Sheet', 'Unknown')}' at row {comp_idx+2}" # +2 for Excel row number
                            )
                        )
                        logger.warning(f"UNIT MISMATCH WARNING: Master row {master_idx} (unit: {master_unit}) vs. Comparison row {comp_idx} (unit: {comp_unit})")

                    # Map DataFrame column names to expected string values
                    column_mapping = {}
                    for i, col_name in enumerate(self.master_dataset.columns):
                        # Map common column names to expected string values
                        if col_name.lower() in ['description', 'desc', 'item']:
                            column_mapping[i] = "DESCRIPTION"
                        elif col_name.lower() in ['quantity', 'qty', 'qty.']:
                            column_mapping[i] = "QUANTITY"
                        elif col_name.lower() in ['unit_price', 'unit price', 'price', 'rate']:
                            column_mapping[i] = "UNIT_PRICE"
                        elif col_name.lower() in ['total_price', 'total price', 'amount', 'total']:
                            column_mapping[i] = "TOTAL_PRICE"
                        elif col_name.lower() in ['code', 'item code', 'ref']:
                            column_mapping[i] = "CODE"
                        elif col_name.lower() in ['unit', 'uom']:
                            column_mapping[i] = "UNIT"
                        elif col_name.lower() in ['manhours', 'ore/u.m.', 'ore', 'man hours']:
                            column_mapping[i] = "MANHOURS"
                        elif col_name.lower() in ['wage', 'euro/hour', 'hourly rate']:
                            column_mapping[i] = "WAGE"
                        else:
                            # Default to DESCRIPTION for unknown columns
                            column_mapping[i] = "DESCRIPTION"
                    
                    # Create comp_row_values based on the comparison data columns
                    comp_row_values = []
                    for col in self.comparison_data.columns:
                        if col in comp_row.index:
                            comp_row_values.append(str(comp_row[col]) if comp_row[col] is not None else '')
                        else:
                            comp_row_values.append('')  # Fill missing columns with empty string
                    
                    # Map DataFrame column names to expected string values for the comparison row
                    column_mapping_for_engine = {}
                    for i, col_name in enumerate(self.comparison_data.columns):
                        if col_name.lower() in ['description', 'desc', 'item']:
                            column_mapping_for_engine[i] = "DESCRIPTION"
                        elif col_name.lower() in ['quantity', 'qty', 'qty.']:
                            column_mapping_for_engine[i] = "QUANTITY"
                        elif col_name.lower() in ['unit_price', 'unit price', 'price', 'rate']:
                            column_mapping_for_engine[i] = "UNIT_PRICE"
                        elif col_name.lower() in ['total_price', 'total price', 'amount', 'total']:
                            column_mapping_for_engine[i] = "TOTAL_PRICE"
                        elif col_name.lower() in ['code', 'item code', 'ref']:
                            column_mapping_for_engine[i] = "CODE"
                        elif col_name.lower() in ['unit', 'uom']:
                            column_mapping_for_engine[i] = "UNIT"
                        elif col_name.lower() in ['manhours', 'ore/u.m.', 'ore', 'man hours']:
                            column_mapping_for_engine[i] = "MANHOURS"
                        elif col_name.lower() in ['wage', 'euro/hour', 'hourly rate']:
                            column_mapping_for_engine[i] = "WAGE"
                        elif col_name.lower() in ['scope']:
                            column_mapping_for_engine[i] = "SCOPE"
                        # Don't map unknown columns to DESCRIPTION - leave them unmapped
                    
                    merge_result = comparison_engine.MERGE(
                        comp_row_values,
                        self.master_dataset,
                        offer_name=offer_name or "ComparisonOffer",
                        column_mapping=column_mapping_for_engine, # Use the new mapping
                        row_index=master_idx
                    )
                    merge_results.append({
                        'type': 'MERGE',
                        'comp_row_index': comp_idx,
                        'master_row_index': master_idx,
                        'result': merge_result
                    })
                else:
                    # ADD: add comp_row to master_dataset
                    # logger.info(f"Row {idx} - ADD decision: instance {n} -> new row (no matching master instance)")
                    if is_row_195:
                        # logger.info(f"Row 195 - ADD: Will add as new row (instance {n} >= {len(master_instances)} master instances)")
                        # logger.info(f"Row 195 - ADD: This is why we're seeing an ADD instead of MERGE!")
                        pass
                    
                    # Map DataFrame column names to expected string values
                    column_mapping = {}
                    for i, col_name in enumerate(self.master_dataset.columns):
                        # Map common column names to expected string values
                        if col_name.lower() in ['description', 'desc', 'item']:
                            column_mapping[i] = "DESCRIPTION"
                        elif col_name.lower() in ['quantity', 'qty', 'qty.']:
                            column_mapping[i] = "QUANTITY"
                        elif col_name.lower() in ['unit_price', 'unit price', 'price', 'rate']:
                            column_mapping[i] = "UNIT_PRICE"
                        elif col_name.lower() in ['total_price', 'total price', 'amount', 'total']:
                            column_mapping[i] = "TOTAL_PRICE"
                        elif col_name.lower() in ['code', 'item code', 'ref']:
                            column_mapping[i] = "CODE"
                        elif col_name.lower() in ['unit', 'uom']:
                            column_mapping[i] = "UNIT"
                        elif col_name.lower() in ['manhours', 'ore/u.m.', 'ore', 'man hours']:
                            column_mapping[i] = "MANHOURS"
                        elif col_name.lower() in ['wage', 'euro/hour', 'hourly rate']:
                            column_mapping[i] = "WAGE"
                        else:
                            # Default to DESCRIPTION for unknown columns
                            column_mapping[i] = "DESCRIPTION"
                    
                    # Create comp_row_values based on the comparison data columns
                    comp_row_values = []
                    for col in self.comparison_data.columns:
                        if col in comp_row.index:
                            comp_row_values.append(str(comp_row[col]) if comp_row[col] is not None else '')
                        else:
                            comp_row_values.append('')  # Fill missing columns with empty string
                    
                    # Use the SAME mapping logic as MERGE operations (lines 744-763)
                    column_mapping_for_engine = {}
                    for i, col_name in enumerate(self.comparison_data.columns):
                        if col_name.lower() in ['description', 'desc', 'item']:
                            column_mapping_for_engine[i] = "DESCRIPTION"
                        elif col_name.lower() in ['quantity', 'qty', 'qty.']:
                            column_mapping_for_engine[i] = "QUANTITY"
                        elif col_name.lower() in ['unit_price', 'unit price', 'price', 'rate']:
                            column_mapping_for_engine[i] = "UNIT_PRICE"
                        elif col_name.lower() in ['total_price', 'total price', 'amount', 'total']:
                            column_mapping_for_engine[i] = "TOTAL_PRICE"
                        elif col_name.lower() in ['code', 'item code', 'ref']:
                            column_mapping_for_engine[i] = "CODE"
                        elif col_name.lower() in ['unit', 'uom']:
                            column_mapping_for_engine[i] = "UNIT"
                        elif col_name.lower() in ['manhours', 'ore/u.m.', 'ore', 'man hours']:
                            column_mapping_for_engine[i] = "MANHOURS"
                        elif col_name.lower() in ['wage', 'euro/hour', 'hourly rate']:
                            column_mapping_for_engine[i] = "WAGE"
                        elif col_name.lower() in ['scope']:
                            column_mapping_for_engine[i] = "SCOPE"
                        # Don't map unknown columns to DESCRIPTION - leave them unmapped
                    
                    logger.warning(f"PROCESS_VALID_ROWS DEBUG: About to call ADD function")
                    logger.warning(f"PROCESS_VALID_ROWS DEBUG: comp_row_values = {comp_row_values}")
                    logger.warning(f"PROCESS_VALID_ROWS DEBUG: column_mapping_for_engine = {column_mapping_for_engine}")
                    logger.error(f"PROCESS_VALID_ROWS DEBUG: comparison_engine type = {type(comparison_engine)}")
                    logger.error(f"PROCESS_VALID_ROWS DEBUG: comparison_engine.ADD method = {comparison_engine.ADD}")
                    
                    try:
                        logger.error("=== ABOUT TO CALL ADD FUNCTION ===")
                        add_result = comparison_engine.ADD(
                            comp_row_values,
                            self.master_dataset,
                            column_mapping=column_mapping_for_engine, # Use the new mapping
                            position=comp_row.get('Position', len(self.master_dataset) + 1),
                            offer_name=offer_name,
                            source_sheet=comp_row.get('Source_Sheet', 'Comparison')  # Get source sheet from comparison data
                        )
                        logger.error("=== ADD FUNCTION RETURNED ===")
                    except Exception as e:
                        logger.error(f"=== EXCEPTION IN ADD FUNCTION CALL: {e} ===")
                        logger.error(f"=== EXCEPTION TYPE: {type(e)} ===")
                        import traceback
                        logger.error(f"=== TRACEBACK: {traceback.format_exc()} ===")
                        raise
                    
                    logger.warning(f"PROCESS_VALID_ROWS DEBUG: ADD result = {add_result}")
                    add_results.append({
                        'type': 'ADD',
                        'comp_row_index': comp_idx,
                        'result': add_result
                    })
        self.merge_results = merge_results
        self.add_results = add_results
        
        # Final logging summary
        logger.info(f"=== PROCESS_VALID_ROWS COMPLETED ===")
        logger.info(f"Total MERGE operations: {len(merge_results)}")
        logger.info(f"Total ADD operations: {len(add_results)}")
        
        if len(add_results) > 0:
            logger.warning(f"WARNING: Found {len(add_results)} ADD operations when processing identical datasets!")
            for add_op in add_results:
                logger.warning(f"ADD operation for comparison row {add_op['comp_row_index']}")
        
        return merge_results + add_results 

    def cleanup_comparison_data(self, recategorize_func=None, numeric_columns=None, category_column='Category'):
        """
        1. Replace empty cells in Unitary_Price or Total_Price with zero values
        2. Collect all rows with empty categories
        3. Call RECATEGORIZATION function for uncategorized rows
        Args:
            recategorize_func: function to call for recategorization (optional)
            numeric_columns: list of columns to treat as numeric (default: ['Unit_Price', 'Total_Price'])
            category_column: name of the category column (default: 'Category')
        Returns:
            List of recategorization results (if any)
        """
        import numpy as np
        if self.comparison_data is None:
            raise ValueError("Comparison data not loaded.")
        if numeric_columns is None:
            numeric_columns = ['Unit_Price', 'Total_Price']
        # 1. Replace empty numeric cells with zero
        for col in numeric_columns:
            if col in self.comparison_data.columns:
                self.comparison_data[col] = self.comparison_data[col].replace([None, '', np.nan], 0)
        # 2. Collect uncategorized rows
        uncategorized_mask = (self.comparison_data[category_column].isnull() |
                              (self.comparison_data[category_column] == '') |
                              (self.comparison_data[category_column].astype(str).str.strip() == ''))
        uncategorized_rows = self.comparison_data[uncategorized_mask]
        recat_results = []
        # 3. Call RECATEGORIZATION for uncategorized rows
        if recategorize_func is not None and not uncategorized_rows.empty:
            for idx, row in uncategorized_rows.iterrows():
                result = recategorize_func(row)
                recat_results.append({'row_index': idx, 'result': result})
        return recat_results 