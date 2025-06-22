"""
Row Classifier for BOQ Tools
Intelligent row classification with data completeness scoring and pattern detection
"""

import logging
import re
from typing import Dict, List, Tuple, Optional, Any, Set
from dataclasses import dataclass
from enum import Enum

from utils.config import get_config, ColumnType

logger = logging.getLogger(__name__)


class RowType(Enum):
    """Types of rows in BOQ sheets"""
    PRIMARY_LINE_ITEM = "primary_line_item"
    HEADER_SECTION_BREAK = "header_section_break"
    SUBTOTAL_ROW = "subtotal_row"
    NOTES_COMMENTS = "notes_comments"
    BLANK_SEPARATOR = "blank_separator"
    INVALID_LINE_ITEM = "invalid_line_item"


@dataclass
class RowClassification:
    """Classification result for a single row"""
    row_index: int
    row_type: RowType
    confidence: float
    reasoning: List[str]
    completeness_score: float
    validation_errors: List[str]
    hierarchical_level: Optional[int]
    section_title: Optional[str]


@dataclass
class ClassificationResult:
    """Result of row classification process"""
    classifications: List[RowClassification]
    summary: Dict[RowType, int]
    overall_quality_score: float
    suggestions: List[str]


class RowClassifier:
    """
    Intelligent row classifier with data completeness scoring and pattern detection
    """
    
    def __init__(self, min_completeness_threshold: float = 0.7):
        """
        Initialize the row classifier
        
        Args:
            min_completeness_threshold: Minimum completeness score for primary line items
        """
        self.config = get_config()
        self.min_completeness_threshold = min_completeness_threshold
        self._setup_patterns()
        
        logger.info("Row Classifier initialized")
    
    def _setup_patterns(self):
        """Setup patterns for row classification"""
        # Subtotal patterns
        self.subtotal_patterns = [
            r'\b(sub)?total\b',
            r'\bsum\b',
            r'\baggregate\b',
            r'\bcombined\b',
            r'\boverall\b',
            r'\bnet\b',
            r'\bgrand\s+total\b',
            r'\bless\b',
            r'\bplus\b',
            r'\badd\b',
            r'\bdeduct\b'
        ]
        
        # Header/section break patterns
        self.header_patterns = [
            r'^[A-Z\s]+$',  # ALL CAPS
            r'^[A-Z][a-z\s]+:$',  # Title case with colon
            r'^[A-Z][a-z\s]+\([^)]+\)$',  # Title with parentheses
            r'^Section\s+\d+',
            r'^Chapter\s+\d+',
            r'^Part\s+\d+',
            r'^Division\s+\d+',
            r'^Subdivision\s+\d+'
        ]
        
        # Hierarchical numbering patterns
        self.hierarchical_patterns = [
            r'^(\d+)\.(\d+)(\.\d+)*$',  # 1.1, 1.1.1, etc.
            r'^([A-Z])\.(\d+)(\.\d+)*$',  # A.1, A.1.1, etc.
            r'^(\d+)-(\d+)(-\d+)*$',  # 1-1, 1-1-1, etc.
            r'^([A-Z])-(\d+)(-\d+)*$',  # A-1, A-1-1, etc.
            r'^(\d+)\)(\d+)(\)\d+)*$',  # 1)1)1, etc.
            r'^([A-Z])\)(\d+)(\)\d+)*$'  # A)1)1, etc.
        ]
        
        # Notes/comment indicators
        self.notes_patterns = [
            r'\bnote\b',
            r'\bcomment\b',
            r'\bremark\b',
            r'\bobservation\b',
            r'\binclude\b',
            r'\bexclude\b',
            r'\bassume\b',
            r'\bpermit\b',
            r'\ballow\b',
            r'\bprovisional\b',
            r'\bprovisional\b',
            r'\bcontingency\b',
            r'\brisk\b',
            r'\bvariation\b'
        ]
        
        # Required columns for line items
        self.required_columns = [
            ColumnType.DESCRIPTION,
            ColumnType.QUANTITY,
            ColumnType.UNIT_PRICE
        ]
        
        # Optional columns that improve completeness
        self.optional_columns = [
            ColumnType.CODE,
            ColumnType.UNIT,
            ColumnType.TOTAL_PRICE
        ]
    
    def classify_rows(self, sheet_data: List[List[str]], 
                     column_mapping: Dict[int, ColumnType]) -> ClassificationResult:
        """
        Classify all rows in a sheet
        
        Args:
            sheet_data: Sheet data as list of rows
            column_mapping: Dictionary mapping column index to ColumnType
            
        Returns:
            ClassificationResult with all row classifications
        """
        logger.info(f"Classifying {len(sheet_data)} rows")
        
        classifications = []
        
        for row_index, row_data in enumerate(sheet_data):
            try:
                classification = self._classify_single_row(row_index, row_data, column_mapping)
                classifications.append(classification)
            except Exception as e:
                logger.warning(f"Error classifying row {row_index}: {e}")
                # Create fallback classification
                classification = RowClassification(
                    row_index=row_index,
                    row_type=RowType.BLANK_SEPARATOR,
                    confidence=0.0,
                    reasoning=[f"Error during classification: {e}"],
                    completeness_score=0.0,
                    validation_errors=[],
                    hierarchical_level=None,
                    section_title=None
                )
                classifications.append(classification)
        
        # Generate summary and suggestions
        summary = self._generate_summary(classifications)
        overall_quality = self._calculate_overall_quality(classifications)
        suggestions = self._generate_suggestions(classifications, summary)
        
        result = ClassificationResult(
            classifications=classifications,
            summary=summary,
            overall_quality_score=overall_quality,
            suggestions=suggestions
        )
        
        logger.info(f"Classification completed: {summary}")
        return result
    
    def _classify_single_row(self, row_index: int, row_data: List[str], 
                           column_mapping: Dict[int, ColumnType]) -> RowClassification:
        """Classify a single row"""
        # Calculate completeness score
        completeness_score = self.calculate_completeness_score(row_data, column_mapping)
        
        # Detect patterns
        is_subtotal = self.detect_subtotal_patterns(row_data)
        is_header = self._detect_header_patterns(row_data)
        is_notes = self._detect_notes_patterns(row_data)
        hierarchical_level = self._detect_hierarchical_level(row_data)
        section_title = self._extract_section_title(row_data)
        
        # Determine row type
        row_type, confidence, reasoning = self._determine_row_type(
            row_data, completeness_score, is_subtotal, is_header, is_notes, hierarchical_level
        )
        
        # Validate if it's a line item
        validation_errors = []
        if row_type == RowType.PRIMARY_LINE_ITEM:
            validation_errors = self.validate_line_item(row_data, column_mapping)
            if validation_errors:
                row_type = RowType.INVALID_LINE_ITEM
                confidence *= 0.5  # Reduce confidence for invalid items
        
        return RowClassification(
            row_index=row_index,
            row_type=row_type,
            confidence=confidence,
            reasoning=reasoning,
            completeness_score=completeness_score,
            validation_errors=validation_errors,
            hierarchical_level=hierarchical_level,
            section_title=section_title
        )
    
    def calculate_completeness_score(self, row_data: List[str], 
                                   column_mapping: Dict[int, ColumnType]) -> float:
        """
        Calculate data completeness score for a row
        
        Args:
            row_data: Row data as list of cell values
            column_mapping: Dictionary mapping column index to ColumnType
            
        Returns:
            Completeness score between 0.0 and 1.0
        """
        if not row_data or not column_mapping:
            return 0.0
        
        # Count required and optional columns
        required_count = 0
        optional_count = 0
        total_required = len(self.required_columns)
        total_optional = len(self.optional_columns)
        
        # Check each column
        for col_idx, col_type in column_mapping.items():
            if col_idx < len(row_data):
                cell_value = row_data[col_idx].strip() if row_data[col_idx] else ""
                
                if col_type in self.required_columns and cell_value:
                    required_count += 1
                elif col_type in self.optional_columns and cell_value:
                    optional_count += 1
        
        # Calculate score with weights
        required_score = required_count / total_required if total_required > 0 else 0.0
        optional_score = optional_count / total_optional if total_optional > 0 else 0.0
        
        # Weight required columns more heavily (70% required, 30% optional)
        completeness_score = (required_score * 0.7) + (optional_score * 0.3)
        
        return min(completeness_score, 1.0)
    
    def detect_subtotal_patterns(self, row_data: List[str]) -> bool:
        """
        Detect if row contains subtotal patterns
        
        Args:
            row_data: Row data as list of cell values
            
        Returns:
            True if subtotal patterns detected
        """
        if not row_data:
            return False
        
        # Check each cell for subtotal patterns
        for cell in row_data:
            if not cell:
                continue
            
            cell_lower = str(cell).lower().strip()
            
            # Check against subtotal patterns
            for pattern in self.subtotal_patterns:
                if re.search(pattern, cell_lower, re.IGNORECASE):
                    return True
        
        return False
    
    def validate_line_item(self, row_data: List[str], 
                          column_mapping: Dict[int, ColumnType]) -> List[str]:
        """
        Validate a line item row
        
        Args:
            row_data: Row data as list of cell values
            column_mapping: Dictionary mapping column index to ColumnType
            
        Returns:
            List of validation errors
        """
        errors = []
        
        if not row_data or not column_mapping:
            errors.append("No data or column mapping provided")
            return errors
        
        # Check for required columns
        for col_type in self.required_columns:
            col_indices = [idx for idx, ct in column_mapping.items() if ct == col_type]
            if not col_indices:
                errors.append(f"Missing required column: {col_type.value}")
                continue
            
            # Check if any of the required columns have data
            has_data = False
            for col_idx in col_indices:
                if col_idx < len(row_data) and row_data[col_idx].strip():
                    has_data = True
                    break
            
            if not has_data:
                errors.append(f"Missing data in required column: {col_type.value}")
        
        # Validate numeric fields
        for col_idx, col_type in column_mapping.items():
            if col_idx >= len(row_data):
                continue
            
            cell_value = row_data[col_idx].strip()
            if not cell_value:
                continue
            
            # Validate quantity
            if col_type == ColumnType.QUANTITY:
                if not self._is_positive_numeric(cell_value):
                    errors.append(f"Invalid quantity: '{cell_value}' (must be positive number)")
            
            # Validate unit price
            elif col_type == ColumnType.UNIT_PRICE:
                if not self._is_positive_numeric(cell_value):
                    errors.append(f"Invalid unit price: '{cell_value}' (must be positive number)")
            
            # Validate total price
            elif col_type == ColumnType.TOTAL_PRICE:
                if not self._is_positive_numeric(cell_value):
                    errors.append(f"Invalid total price: '{cell_value}' (must be positive number)")
        
        return errors
    
    def get_row_confidence(self, row_data: List[str], 
                          classification: RowClassification) -> float:
        """
        Calculate confidence score for a row classification
        
        Args:
            row_data: Row data as list of cell values
            classification: Row classification result
            
        Returns:
            Confidence score between 0.0 and 1.0
        """
        confidence = classification.confidence
        
        # Adjust confidence based on data quality
        if classification.row_type == RowType.PRIMARY_LINE_ITEM:
            # Higher completeness = higher confidence
            confidence *= (0.5 + classification.completeness_score * 0.5)
            
            # Penalize for validation errors
            if classification.validation_errors:
                confidence *= 0.7
        
        elif classification.row_type == RowType.SUBTOTAL_ROW:
            # Check for numeric values in expected columns
            numeric_count = sum(1 for cell in row_data if self._is_numeric(cell))
            if numeric_count > 0:
                confidence *= 1.2  # Boost confidence
            else:
                confidence *= 0.8  # Reduce confidence
        
        elif classification.row_type == RowType.HEADER_SECTION_BREAK:
            # Check for hierarchical numbering
            if classification.hierarchical_level is not None:
                confidence *= 1.1  # Boost confidence
        
        elif classification.row_type == RowType.NOTES_COMMENTS:
            # Check for descriptive text
            text_cells = sum(1 for cell in row_data if cell and len(cell) > 10)
            if text_cells > 0:
                confidence *= 1.1  # Boost confidence
        
        elif classification.row_type == RowType.BLANK_SEPARATOR:
            # Check if truly blank
            non_empty_cells = sum(1 for cell in row_data if cell and cell.strip())
            if non_empty_cells == 0:
                confidence *= 1.2  # Boost confidence
            else:
                confidence *= 0.8  # Reduce confidence
        
        return min(confidence, 1.0)
    
    def _detect_header_patterns(self, row_data: List[str]) -> bool:
        """Detect header/section break patterns"""
        if not row_data:
            return False
        
        # Check if any cell matches header patterns
        for cell in row_data:
            if not cell:
                continue
            
            cell_str = str(cell).strip()
            
            for pattern in self.header_patterns:
                if re.match(pattern, cell_str, re.IGNORECASE):
                    return True
        
        return False
    
    def _detect_notes_patterns(self, row_data: List[str]) -> bool:
        """Detect notes/comment patterns"""
        if not row_data:
            return False
        
        # Check if any cell contains notes keywords
        for cell in row_data:
            if not cell:
                continue
            
            cell_lower = str(cell).lower().strip()
            
            for pattern in self.notes_patterns:
                if re.search(pattern, cell_lower, re.IGNORECASE):
                    return True
        
        return False
    
    def _detect_hierarchical_level(self, row_data: List[str]) -> Optional[int]:
        """Detect hierarchical numbering level"""
        if not row_data:
            return None
        
        # Check first few cells for hierarchical patterns
        for cell in row_data[:3]:  # Check first 3 cells
            if not cell:
                continue
            
            cell_str = str(cell).strip()
            
            for pattern in self.hierarchical_patterns:
                match = re.match(pattern, cell_str)
                if match:
                    # Count the number of levels (dots, dashes, parentheses)
                    levels = len(re.findall(r'[.\-)]', cell_str))
                    return levels + 1  # Add 1 for base level
        
        return None
    
    def _extract_section_title(self, row_data: List[str]) -> Optional[str]:
        """Extract section title from row data"""
        if not row_data:
            return None
        
        # Look for the most likely title cell
        for cell in row_data:
            if not cell:
                continue
            
            cell_str = str(cell).strip()
            
            # Check if it looks like a title
            if (len(cell_str) > 3 and 
                (cell_str.isupper() or 
                 re.match(r'^[A-Z][a-z\s]+', cell_str) or
                 re.match(r'^[A-Z][a-z\s]+:', cell_str))):
                return cell_str
        
        return None
    
    def _determine_row_type(self, row_data: List[str], completeness_score: float,
                           is_subtotal: bool, is_header: bool, is_notes: bool,
                           hierarchical_level: Optional[int]) -> Tuple[RowType, float, List[str]]:
        """Determine the type of a row based on patterns and completeness"""
        reasoning = []
        confidence = 0.0
        
        # Check for blank/separator rows first
        non_empty_cells = sum(1 for cell in row_data if cell and cell.strip())
        if non_empty_cells == 0:
            reasoning.append("No non-empty cells found")
            return RowType.BLANK_SEPARATOR, 0.9, reasoning
        
        # Check for subtotal patterns
        if is_subtotal:
            reasoning.append("Subtotal patterns detected")
            confidence = 0.8
            return RowType.SUBTOTAL_ROW, confidence, reasoning
        
        # Check for header patterns
        if is_header:
            reasoning.append("Header/section break patterns detected")
            if hierarchical_level is not None:
                reasoning.append(f"Hierarchical level {hierarchical_level} detected")
            confidence = 0.85
            return RowType.HEADER_SECTION_BREAK, confidence, reasoning
        
        # Check for notes patterns
        if is_notes:
            reasoning.append("Notes/comment patterns detected")
            confidence = 0.75
            return RowType.NOTES_COMMENTS, confidence, reasoning
        
        # Check for primary line items
        if completeness_score >= self.min_completeness_threshold:
            reasoning.append(f"High completeness score: {completeness_score:.2f}")
            confidence = 0.8 + (completeness_score * 0.2)
            return RowType.PRIMARY_LINE_ITEM, confidence, reasoning
        
        # Check for partial line items
        elif completeness_score > 0.3:
            reasoning.append(f"Moderate completeness score: {completeness_score:.2f}")
            confidence = 0.6
            return RowType.PRIMARY_LINE_ITEM, confidence, reasoning
        
        # Default to notes/comments for low completeness
        else:
            reasoning.append(f"Low completeness score: {completeness_score:.2f}")
            confidence = 0.5
            return RowType.NOTES_COMMENTS, confidence, reasoning
    
    def _is_positive_numeric(self, value: str) -> bool:
        """Check if value is a positive number"""
        try:
            # Remove currency symbols and commas
            clean_value = re.sub(r'[\$€£¥₹,]', '', value)
            num = float(clean_value)
            return num > 0
        except (ValueError, TypeError):
            return False
    
    def _is_numeric(self, value: str) -> bool:
        """Check if value is numeric"""
        try:
            # Remove currency symbols and commas
            clean_value = re.sub(r'[\$€£¥₹,]', '', value)
            float(clean_value)
            return True
        except (ValueError, TypeError):
            return False
    
    def _generate_summary(self, classifications: List[RowClassification]) -> Dict[RowType, int]:
        """Generate summary of row types"""
        summary = {row_type: 0 for row_type in RowType}
        
        for classification in classifications:
            summary[classification.row_type] += 1
        
        return summary
    
    def _calculate_overall_quality(self, classifications: List[RowClassification]) -> float:
        """Calculate overall quality score"""
        if not classifications:
            return 0.0
        
        # Calculate weighted average confidence
        total_weight = 0.0
        weighted_sum = 0.0
        
        for classification in classifications:
            # Weight by row type importance
            if classification.row_type == RowType.PRIMARY_LINE_ITEM:
                weight = 2.0  # Most important
            elif classification.row_type == RowType.SUBTOTAL_ROW:
                weight = 1.5  # Important for structure
            elif classification.row_type == RowType.HEADER_SECTION_BREAK:
                weight = 1.2  # Important for organization
            else:
                weight = 1.0  # Standard weight
            
            weighted_sum += classification.confidence * weight
            total_weight += weight
        
        return weighted_sum / total_weight if total_weight > 0 else 0.0
    
    def _generate_suggestions(self, classifications: List[RowClassification], 
                            summary: Dict[RowType, int]) -> List[str]:
        """Generate suggestions for improving the sheet"""
        suggestions = []
        
        # Check for missing line items
        if summary[RowType.PRIMARY_LINE_ITEM] == 0:
            suggestions.append("No primary line items found - check if headers are correctly identified")
        
        # Check for too many invalid items
        invalid_count = summary[RowType.INVALID_LINE_ITEM]
        total_items = summary[RowType.PRIMARY_LINE_ITEM] + invalid_count
        if total_items > 0 and invalid_count / total_items > 0.3:
            suggestions.append(f"High number of invalid line items ({invalid_count}/{total_items}) - check data quality")
        
        # Check for missing subtotals
        if summary[RowType.PRIMARY_LINE_ITEM] > 10 and summary[RowType.SUBTOTAL_ROW] == 0:
            suggestions.append("Many line items but no subtotals found - consider adding summary rows")
        
        # Check for too many blank rows
        blank_count = summary[RowType.BLANK_SEPARATOR]
        total_rows = len(classifications)
        if blank_count / total_rows > 0.2:
            suggestions.append(f"High number of blank rows ({blank_count}/{total_rows}) - consider cleaning up")
        
        # Check for hierarchical structure
        hierarchical_rows = [c for c in classifications if c.hierarchical_level is not None]
        if hierarchical_rows:
            levels = set(c.hierarchical_level for c in hierarchical_rows)
            if len(levels) > 3:
                suggestions.append(f"Complex hierarchical structure with {len(levels)} levels - ensure consistency")
        
        return suggestions


# Convenience function for quick row classification
def classify_rows_quick(sheet_data: List[List[str]], 
                       column_mapping: Dict[int, str]) -> Dict[int, str]:
    """
    Quick row classification
    
    Args:
        sheet_data: Sheet data as list of rows
        column_mapping: Dictionary mapping column index to column type string
        
    Returns:
        Dictionary mapping row index to row type
    """
    # Convert string column types to ColumnType enum
    enum_mapping = {}
    for col_idx, col_type_str in column_mapping.items():
        try:
            enum_mapping[col_idx] = ColumnType(col_type_str)
        except ValueError:
            logger.warning(f"Unknown column type: {col_type_str}")
    
    classifier = RowClassifier()
    result = classifier.classify_rows(sheet_data, enum_mapping)
    
    return {classification.row_index: classification.row_type.value 
            for classification in result.classifications} 