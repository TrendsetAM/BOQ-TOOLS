"""
Instance Matcher for BOQ Tools
Handles finding and matching instances of the same description across different datasets
"""

import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

logger = logging.getLogger(__name__)


class DatasetType(Enum):
    """Types of datasets for comparison"""
    COMPARISON_BOQ = "comparison_boq"
    DATASET = "dataset"


@dataclass
class RowInstance:
    """Represents a row instance with position and data"""
    position: int
    row_data: List[str]
    sheet_name: str
    row_index: int
    description: str
    instance_number: int


class InstanceMatcher:
    """
    Matches instances of the same description across different datasets
    """
    
    def __init__(self):
        """Initialize the instance matcher"""
        logger.info("Instance Matcher initialized")
    
    def LIST_INSTANCES(self, current_position: int, description: str, 
                       dataset_type: DatasetType, all_rows: List[RowInstance]) -> List[RowInstance]:
        """
        LIST_INSTANCES Function: Creates a list of rows having the same exact description, ordered by Position
        
        Args:
            current_position: Position of the current row
            description: Description to match
            dataset_type: Whether to use Comparison BoQ or Dataset
            all_rows: List of all RowInstance objects
            
        Returns:
            List of RowInstance objects with the same description, ordered by Position
        """
        try:
            # Filter rows by description (case-insensitive)
            matching_rows = []
            
            for row_instance in all_rows:
                if (row_instance.description.lower().strip() == description.lower().strip()):
                    matching_rows.append(row_instance)
            
            # Sort by Position to ensure proper ordering
            matching_rows.sort(key=lambda x: x.position)
            
            logger.info(f"Found {len(matching_rows)} instances of '{description}' in {dataset_type.value}")
            
            return matching_rows
            
        except Exception as e:
            logger.error(f"Error in LIST_INSTANCES: {e}")
            return []
    
    def get_comparison_instances(self, current_position: int, description: str, 
                                comparison_rows: List[RowInstance]) -> List[RowInstance]:
        """
        Get instances from Comparison BoQ dataset
        
        Args:
            current_position: Position of the current row
            description: Description to match
            comparison_rows: List of RowInstance objects from Comparison BoQ
            
        Returns:
            List of RowInstance objects with the same description from Comparison BoQ
        """
        return self.LIST_INSTANCES(current_position, description, DatasetType.COMPARISON_BOQ, comparison_rows)
    
    def get_dataset_instances(self, current_position: int, description: str, 
                             dataset_rows: List[RowInstance]) -> List[RowInstance]:
        """
        Get instances from Dataset
        
        Args:
            current_position: Position of the current row
            description: Description to match
            dataset_rows: List of RowInstance objects from Dataset
            
        Returns:
            List of RowInstance objects with the same description from Dataset
        """
        return self.LIST_INSTANCES(current_position, description, DatasetType.DATASET, dataset_rows)
    
    def validate_instance_count(self, comparison_instances: List[RowInstance], 
                               dataset_instances: List[RowInstance], description: str) -> bool:
        """
        Validate that the correct number of instances were found
        
        Args:
            comparison_instances: List of instances from Comparison BoQ
            dataset_instances: List of instances from Dataset
            description: Description being matched
            
        Returns:
            True if validation passes, False otherwise
        """
        try:
            # Log instance counts for debugging
            logger.info(f"Instance validation for '{description}':")
            logger.info(f"  Comparison BoQ instances: {len(comparison_instances)}")
            logger.info(f"  Dataset instances: {len(dataset_instances)}")
            
            # Basic validation: ensure we have at least one instance in each dataset
            if len(comparison_instances) == 0:
                logger.warning(f"No instances found in Comparison BoQ for '{description}'")
                return False
            
            if len(dataset_instances) == 0:
                logger.warning(f"No instances found in Dataset for '{description}'")
                return False
            
            # Additional validation could be added here based on business rules
            # For example, checking if the number of instances makes sense
            
            return True
            
        except Exception as e:
            logger.error(f"Error in instance validation: {e}")
            return False
    
    def create_row_instances_from_data(self, sheet_data: List[List[str]], 
                                      column_mapping: Dict[int, Any],
                                      sheet_name: str, 
                                      position_offset: int = 0) -> List[RowInstance]:
        """
        Create RowInstance objects from sheet data
        
        Args:
            sheet_data: Sheet data as list of rows
            column_mapping: Dictionary mapping column index to ColumnType
            sheet_name: Name of the sheet
            position_offset: Offset to add to position numbers
            
        Returns:
            List of RowInstance objects
        """
        instances = []
        
        for row_index, row_data in enumerate(sheet_data):
            try:
                # Extract description from row data
                description = ""
                for col_idx, col_type in column_mapping.items():
                    if col_idx < len(row_data) and col_type == "DESCRIPTION":
                        description = row_data[col_idx].strip() if row_data[col_idx] else ""
                        break
                
                if description:  # Only create instance if description exists
                    # Calculate position (1-based Excel row number + offset)
                    position = position_offset + row_index + 1
                    
                    # For now, use row index as instance number
                    # In a full implementation, this would be calculated based on description frequency
                    instance_number = row_index + 1
                    
                    instance = RowInstance(
                        position=position,
                        row_data=row_data,
                        sheet_name=sheet_name,
                        row_index=row_index,
                        description=description,
                        instance_number=instance_number
                    )
                    
                    instances.append(instance)
                    
            except Exception as e:
                logger.error(f"Error creating row instance for row {row_index}: {e}")
                continue
        
        return instances 