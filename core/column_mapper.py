"""
Column Mapper for BOQ Tools
Intelligent column mapping with header detection and confidence scoring
"""

import logging
import re
from typing import Dict, List, Tuple, Optional, Any, Set
from dataclasses import dataclass
from enum import Enum

from utils.config import get_config, ColumnType

logger = logging.getLogger(__name__)


class HeaderDetectionMethod(Enum):
    """Methods for header row detection"""
    KEYWORD_MATCH = "keyword_match"
    DATA_TYPE_PATTERN = "data_type_pattern"
    POSITIONAL_LOGIC = "positional_logic"
    MERGED_CELLS = "merged_cells"


@dataclass
class HeaderRowInfo:
    """Information about detected header row"""
    row_index: int
    confidence: float
    method: HeaderDetectionMethod
    reasoning: List[str]
    headers: List[str]
    is_merged: bool


@dataclass
class ColumnMapping:
    """Column mapping information"""
    column_index: int
    original_header: str
    normalized_header: str
    mapped_type: ColumnType
    confidence: float
    alternatives: List[Tuple[ColumnType, float]]
    reasoning: List[str]


@dataclass
class MappingResult:
    """Result of column mapping process"""
    header_row: HeaderRowInfo
    mappings: List[ColumnMapping]
    overall_confidence: float
    unmapped_columns: List[int]
    suggestions: List[str]


class ColumnMapper:
    """
    Intelligent column mapper with header detection and confidence scoring
    """
    
    def __init__(self, max_header_rows: int = 10):
        """
        Initialize the column mapper
        
        Args:
            max_header_rows: Maximum number of rows to check for headers
        """
        self.config = get_config()
        self.max_header_rows = max_header_rows
        self._setup_patterns()
        
        logger.info("Column Mapper initialized")
    
    def _setup_patterns(self):
        """Setup patterns for header detection and normalization"""
        # Currency symbols and patterns
        self.currency_patterns = [
            r'[\$€£¥₹]',  # Currency symbols
            r'price|cost|amount|value|total',  # Financial terms
            r'rate|unit.?price|price.?per.?unit'  # Rate terms
        ]
        
        # Data type indicators
        self.data_type_patterns = {
            'numeric': [
                r'qty|quantity|number|count|nos|units',
                r'amount|price|cost|value|total|sum',
                r'rate|unit.?price|price.?per.?unit',
                r'percentage|%|percent'
            ],
            'text': [
                r'description|desc|item|work|activity|task',
                r'remarks|notes|comments|observation',
                r'type|category|class|classification'
            ],
            'code': [
                r'code|ref|reference|item.?no|item.?number',
                r'boq.?code|schedule.?no|item.?code'
            ],
            'unit': [
                r'unit|units|uom|measurement|measure',
                r'm[²³]|sq\.?m|cu\.?m|kg|ton|l|gal'
            ]
        }
        
        # Positional logic patterns
        self.positional_patterns = {
            'left_columns': ['code', 'description', 'item'],
            'right_columns': ['total', 'amount', 'sum', 'value'],
            'middle_columns': ['quantity', 'unit', 'rate', 'price']
        }
        
        # Symbols to remove during normalization
        self.symbols_to_remove = r'[^\w\s]'
        
        # Common header variations
        self.header_variations = {
            'description': ['desc', 'item', 'work', 'activity', 'task', 'scope'],
            'quantity': ['qty', 'qty.', 'quantity', 'no', 'number', 'count'],
            'unit_price': ['rate', 'price', 'unit price', 'price per unit', 'unit rate'],
            'total': ['amount', 'total', 'sum', 'value', 'cost'],
            'unit': ['unit', 'units', 'uom', 'measurement'],
            'code': ['code', 'ref', 'reference', 'item no', 'item number']
        }
    
    def find_header_row(self, sheet_data: List[List[str]]) -> HeaderRowInfo:
        """
        Find the header row in sheet data
        
        Args:
            sheet_data: Sheet data as list of rows
            
        Returns:
            HeaderRowInfo with header row details
        """
        logger.debug(f"Searching for header row in {len(sheet_data)} rows")
        
        best_header = None
        best_score = 0.0
        
        # Check first few rows for headers
        search_rows = min(self.max_header_rows, len(sheet_data))
        
        for row_index in range(search_rows):
            row = sheet_data[row_index]
            if not row:
                continue
            
            # Try different detection methods
            methods = [
                self._detect_by_keywords,
                self._detect_by_data_patterns,
                self._detect_by_positional_logic,
                self._detect_by_merged_cells
            ]
            
            for method in methods:
                try:
                    result = method(row, row_index, sheet_data)
                    if result and result.confidence > best_score:
                        best_score = result.confidence
                        best_header = result
                except Exception as e:
                    logger.warning(f"Error in header detection method {method.__name__}: {e}")
        
        if not best_header:
            # Fallback: use first non-empty row
            for row_index in range(search_rows):
                row = sheet_data[row_index]
                if row and any(cell.strip() for cell in row):
                    best_header = HeaderRowInfo(
                        row_index=row_index,
                        confidence=0.1,
                        method=HeaderDetectionMethod.KEYWORD_MATCH,
                        reasoning=["Fallback: first non-empty row"],
                        headers=row,
                        is_merged=False
                    )
                    break
        
        if best_header:
            logger.info(f"Header row found at index {best_header.row_index} "
                       f"with confidence {best_header.confidence:.2f}")
        else:
            logger.warning("No header row found")
            best_header = HeaderRowInfo(
                row_index=0,
                confidence=0.0,
                method=HeaderDetectionMethod.KEYWORD_MATCH,
                reasoning=["No header row detected"],
                headers=[],
                is_merged=False
            )
        
        return best_header
    
    def _detect_by_keywords(self, row: List[str], row_index: int, 
                           sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """Detect header row by keyword matching"""
        if not row:
            return None
        
        score = 0.0
        matches = []
        
        # Check each cell for header keywords
        for cell in row:
            if not cell:
                continue
            
            cell_lower = str(cell).lower().strip()
            
            # Check against all column type keywords
            for col_type in self.config.get_all_column_types():
                mapping = self.config.get_column_mapping(col_type)
                if mapping:
                    for keyword in mapping.keywords:
                        if keyword.lower() in cell_lower:
                            score += mapping.weight
                            matches.append(f"'{cell}' matches {col_type.value}")
        
        # Normalize score
        score = min(score / max(1, len(row)), 1.0)
        
        if score > 0.3:  # Threshold for keyword detection
            return HeaderRowInfo(
                row_index=row_index,
                confidence=score,
                method=HeaderDetectionMethod.KEYWORD_MATCH,
                reasoning=[f"Keyword matches: {', '.join(matches[:3])}"],
                headers=row,
                is_merged=False
            )
        
        return None
    
    def _detect_by_data_patterns(self, row: List[str], row_index: int,
                                sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """Detect header row by analyzing data patterns in subsequent rows"""
        if not row or row_index >= len(sheet_data) - 1:
            return None
        
        # Analyze next few rows for data patterns
        data_rows = sheet_data[row_index + 1:row_index + 4]
        if not data_rows:
            return None
        
        score = 0.0
        reasoning = []
        
        # Check if this row has text while next rows have mixed data types
        text_cells = sum(1 for cell in row if cell and not self._is_numeric(cell))
        if text_cells > len(row) * 0.7:  # Mostly text
            score += 0.3
            reasoning.append("Row contains mostly text")
        
        # Check if subsequent rows have numeric data
        numeric_columns = 0
        for data_row in data_rows:
            if data_row:
                numeric_cells = sum(1 for cell in data_row if self._is_numeric(cell))
                if numeric_cells > 0:
                    numeric_columns += 1
        
        if numeric_columns > 0:
            score += 0.4
            reasoning.append(f"Subsequent rows contain numeric data in {numeric_columns} rows")
        
        # Check for currency patterns in subsequent rows
        currency_columns = 0
        for data_row in data_rows:
            if data_row:
                for cell in data_row:
                    if cell and any(re.search(pattern, str(cell), re.IGNORECASE) 
                                  for pattern in self.currency_patterns):
                        currency_columns += 1
                        break
        
        if currency_columns > 0:
            score += 0.3
            reasoning.append(f"Currency patterns detected in {currency_columns} columns")
        
        if score > 0.5:
            return HeaderRowInfo(
                row_index=row_index,
                confidence=score,
                method=HeaderDetectionMethod.DATA_TYPE_PATTERN,
                reasoning=reasoning,
                headers=row,
                is_merged=False
            )
        
        return None
    
    def _detect_by_positional_logic(self, row: List[str], row_index: int,
                                   sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """Detect header row using positional logic"""
        if not row:
            return None
        
        score = 0.0
        reasoning = []
        
        # Check for typical header positioning
        if len(row) >= 3:
            # Left column should have description-like content
            left_cell = str(row[0]).lower() if row[0] else ""
            if any(keyword in left_cell for keyword in self.positional_patterns['left_columns']):
                score += 0.3
                reasoning.append("Left column contains description-like content")
            
            # Right column should have total-like content
            right_cell = str(row[-1]).lower() if row[-1] else ""
            if any(keyword in right_cell for keyword in self.positional_patterns['right_columns']):
                score += 0.3
                reasoning.append("Right column contains total-like content")
            
            # Middle columns should have quantity/unit content
            middle_cells = [str(cell).lower() for cell in row[1:-1] if cell]
            middle_matches = sum(1 for cell in middle_cells 
                               for keyword in self.positional_patterns['middle_columns']
                               if keyword in cell)
            if middle_matches > 0:
                score += 0.2
                reasoning.append(f"Middle columns contain quantity/unit content ({middle_matches} matches)")
        
        if score > 0.4:
            return HeaderRowInfo(
                row_index=row_index,
                confidence=score,
                method=HeaderDetectionMethod.POSITIONAL_LOGIC,
                reasoning=reasoning,
                headers=row,
                is_merged=False
            )
        
        return None
    
    def _detect_by_merged_cells(self, row: List[str], row_index: int,
                               sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """Detect header row by looking for merged cells patterns"""
        if not row:
            return None
        
        # Check for merged cell indicators (empty cells between content)
        empty_cells = sum(1 for cell in row if not cell or not str(cell).strip())
        total_cells = len(row)
        
        if empty_cells > 0 and empty_cells < total_cells:
            # Check if this looks like a merged header
            score = 0.3
            reasoning = [f"Potential merged cells: {empty_cells}/{total_cells} empty cells"]
            
            # Check if next row has more detailed headers
            if row_index + 1 < len(sheet_data):
                next_row = sheet_data[row_index + 1]
                if next_row and len(next_row) > len([c for c in row if c]):
                    score += 0.2
                    reasoning.append("Next row has more detailed headers")
            
            if score > 0.4:
                return HeaderRowInfo(
                    row_index=row_index,
                    confidence=score,
                    method=HeaderDetectionMethod.MERGED_CELLS,
                    reasoning=reasoning,
                    headers=row,
                    is_merged=True
                )
        
        return None
    
    def normalize_header_text(self, headers: List[str]) -> List[str]:
        """
        Normalize header text for better matching
        
        Args:
            headers: List of header strings
            
        Returns:
            List of normalized header strings
        """
        normalized = []
        
        for header in headers:
            if not header:
                normalized.append("")
                continue
            
            # Convert to string and normalize
            header_str = str(header).strip()
            
            # Convert to lowercase
            header_str = header_str.lower()
            
            # Remove common symbols but keep important ones
            header_str = re.sub(r'[^\w\s\-\.]', ' ', header_str)
            
            # Normalize whitespace
            header_str = re.sub(r'\s+', ' ', header_str).strip()
            
            # Handle common variations
            for standard, variations in self.header_variations.items():
                if header_str in variations:
                    header_str = standard
                    break
            
            normalized.append(header_str)
        
        return normalized
    
    def map_columns_to_types(self, headers: List[str]) -> List[ColumnMapping]:
        """
        Map columns to BOQ types using configuration
        
        Args:
            headers: List of header strings
            
        Returns:
            List of ColumnMapping objects
        """
        logger.debug(f"Mapping {len(headers)} columns to BOQ types")
        
        mappings = []
        normalized_headers = self.normalize_header_text(headers)
        
        for col_idx, (original_header, normalized_header) in enumerate(zip(headers, normalized_headers)):
            if not normalized_header:
                continue
            
            # Find best match for this column
            best_type, best_score, alternatives = self._find_best_column_match(normalized_header, col_idx, headers)
            
            if best_type:
                reasoning = self._generate_mapping_reasoning(
                    original_header, normalized_header, best_type, best_score, col_idx, headers
                )
                
                mapping = ColumnMapping(
                    column_index=col_idx,
                    original_header=original_header,
                    normalized_header=normalized_header,
                    mapped_type=best_type,
                    confidence=best_score,
                    alternatives=alternatives,
                    reasoning=reasoning
                )
                
                mappings.append(mapping)
                
                logger.debug(f"Column {col_idx} '{original_header}' -> {best_type.value} "
                           f"(confidence: {best_score:.2f})")
        
        return mappings
    
    def _find_best_column_match(self, normalized_header: str, col_idx: int, 
                               all_headers: List[str]) -> Tuple[Optional[ColumnType], float, List[Tuple[ColumnType, float]]]:
        """Find the best column type match for a header"""
        best_type = None
        best_score = 0.0
        alternatives = []
        
        # Check each column type
        for col_type in self.config.get_all_column_types():
            mapping = self.config.get_column_mapping(col_type)
            if not mapping:
                continue
            
            # Calculate score for this type
            score = self._calculate_header_score(normalized_header, mapping, col_idx, all_headers)
            
            if score > best_score:
                best_score = score
                best_type = col_type
            
            # Store alternatives above threshold
            if score > 0.3:
                alternatives.append((col_type, score))
        
        # Sort alternatives by score
        alternatives.sort(key=lambda x: x[1], reverse=True)
        
        return best_type, best_score, alternatives
    
    def _calculate_header_score(self, normalized_header: str, mapping: Any, 
                               col_idx: int, all_headers: List[str]) -> float:
        """Calculate score for header matching a column type"""
        score = 0.0
        
        # Direct keyword matching
        for keyword in mapping.keywords:
            keyword_lower = keyword.lower()
            if keyword_lower in normalized_header or normalized_header in keyword_lower:
                score += mapping.weight
                break
        
        # Positional scoring
        score += self._calculate_positional_score(col_idx, mapping, all_headers)
        
        # Context scoring (check neighboring columns)
        score += self._calculate_context_score(col_idx, mapping, all_headers)
        
        return min(score, 1.0)
    
    def _calculate_positional_score(self, col_idx: int, mapping: Any, 
                                   all_headers: List[str]) -> float:
        """Calculate score based on column position"""
        score = 0.0
        total_cols = len(all_headers)
        
        if total_cols == 0:
            return score
        
        # Normalize position (0-1)
        position = col_idx / (total_cols - 1) if total_cols > 1 else 0.5
        
        # Position preferences for different column types
        if mapping.required:
            if "description" in str(mapping).lower():
                # Description typically on the left
                if position < 0.3:
                    score += 0.2
            elif "total" in str(mapping).lower() or "amount" in str(mapping).lower():
                # Totals typically on the right
                if position > 0.7:
                    score += 0.2
            elif "quantity" in str(mapping).lower():
                # Quantity typically in the middle
                if 0.2 < position < 0.8:
                    score += 0.1
        
        return score
    
    def _calculate_context_score(self, col_idx: int, mapping: Any, 
                                all_headers: List[str]) -> float:
        """Calculate score based on neighboring columns"""
        score = 0.0
        
        # Check left neighbor
        if col_idx > 0:
            left_header = all_headers[col_idx - 1].lower() if all_headers[col_idx - 1] else ""
            if "description" in left_header and "quantity" in str(mapping).lower():
                score += 0.1  # Quantity often follows description
            elif "quantity" in left_header and "unit" in str(mapping).lower():
                score += 0.1  # Unit often follows quantity
        
        # Check right neighbor
        if col_idx < len(all_headers) - 1:
            right_header = all_headers[col_idx + 1].lower() if all_headers[col_idx + 1] else ""
            if "total" in right_header and "rate" in str(mapping).lower():
                score += 0.1  # Rate often precedes total
        
        return score
    
    def _generate_mapping_reasoning(self, original_header: str, normalized_header: str,
                                   col_type: ColumnType, score: float, col_idx: int,
                                   all_headers: List[str]) -> List[str]:
        """Generate reasoning for column mapping"""
        reasoning = []
        
        reasoning.append(f"Column {col_idx + 1}: '{original_header}' -> {col_type.value}")
        reasoning.append(f"Confidence: {score:.2f}")
        
        # Add keyword match reasoning
        mapping = self.config.get_column_mapping(col_type)
        if mapping:
            for keyword in mapping.keywords:
                if keyword.lower() in normalized_header:
                    reasoning.append(f"Keyword match: '{keyword}'")
                    break
        
        # Add positional reasoning
        total_cols = len(all_headers)
        if total_cols > 1:
            position = col_idx / (total_cols - 1)
            reasoning.append(f"Position: {position:.1f} ({col_idx + 1}/{total_cols})")
        
        return reasoning
    
    def calculate_mapping_confidence(self, mappings: List[ColumnMapping]) -> float:
        """
        Calculate overall confidence for column mappings
        
        Args:
            mappings: List of column mappings
            
        Returns:
            Overall confidence score
        """
        if not mappings:
            return 0.0
        
        # Calculate weighted average confidence
        total_weight = 0.0
        weighted_sum = 0.0
        
        for mapping in mappings:
            # Weight by column importance (required columns get higher weight)
            col_mapping = self.config.get_column_mapping(mapping.mapped_type)
            weight = col_mapping.weight if col_mapping else 1.0
            
            weighted_sum += mapping.confidence * weight
            total_weight += weight
        
        overall_confidence = weighted_sum / total_weight if total_weight > 0 else 0.0
        
        logger.debug(f"Overall mapping confidence: {overall_confidence:.2f}")
        return overall_confidence
    
    def get_alternative_mappings(self, headers: List[str]) -> Dict[int, List[Tuple[ColumnType, float]]]:
        """
        Get alternative mappings for ambiguous columns
        
        Args:
            headers: List of header strings
            
        Returns:
            Dictionary mapping column index to alternative mappings
        """
        alternatives = {}
        normalized_headers = self.normalize_header_text(headers)
        
        for col_idx, (original_header, normalized_header) in enumerate(zip(headers, normalized_headers)):
            if not normalized_header:
                continue
            
            column_alternatives = []
            
            # Check all column types
            for col_type in self.config.get_all_column_types():
                mapping = self.config.get_column_mapping(col_type)
                if not mapping:
                    continue
                
                score = self._calculate_header_score(normalized_header, mapping, col_idx, headers)
                
                # Include alternatives with reasonable scores
                if score > 0.2:
                    column_alternatives.append((col_type, score))
            
            # Sort by score and keep top alternatives
            column_alternatives.sort(key=lambda x: x[1], reverse=True)
            if column_alternatives:
                alternatives[col_idx] = column_alternatives[:3]  # Top 3 alternatives
        
        return alternatives
    
    def _is_numeric(self, cell: str) -> bool:
        """Check if a cell contains numeric data"""
        if not cell:
            return False
        
        cell_str = str(cell).strip()
        
        # Check for currency symbols
        cell_str = re.sub(r'[\$€£¥₹,]', '', cell_str)
        
        # Check for percentage
        if cell_str.endswith('%'):
            cell_str = cell_str[:-1]
        
        # Check for units
        cell_str = re.sub(r'\s*(m[²³]|sq\.?m|cu\.?m|kg|ton|l|gal)$', '', cell_str, flags=re.IGNORECASE)
        
        # Check if numeric
        try:
            float(cell_str)
            return True
        except ValueError:
            return False
    
    def process_sheet_mapping(self, sheet_data: List[List[str]]) -> MappingResult:
        """
        Complete sheet mapping process
        
        Args:
            sheet_data: Sheet data as list of rows
            
        Returns:
            MappingResult with complete mapping information
        """
        logger.info(f"Processing sheet mapping for {len(sheet_data)} rows")
        
        # Find header row
        header_info = self.find_header_row(sheet_data)
        
        if not header_info or not header_info.headers:
            return MappingResult(
                header_row=header_info,
                mappings=[],
                overall_confidence=0.0,
                unmapped_columns=list(range(len(sheet_data[0]) if sheet_data else 0)),
                suggestions=["No headers found"]
            )
        
        # Map columns to types
        mappings = self.map_columns_to_types(header_info.headers)
        
        # Calculate overall confidence
        overall_confidence = self.calculate_mapping_confidence(mappings)
        
        # Find unmapped columns
        mapped_indices = {m.column_index for m in mappings}
        unmapped_columns = [i for i in range(len(header_info.headers)) if i not in mapped_indices]
        
        # Generate suggestions
        suggestions = self._generate_mapping_suggestions(mappings, unmapped_columns, header_info.headers)
        
        result = MappingResult(
            header_row=header_info,
            mappings=mappings,
            overall_confidence=overall_confidence,
            unmapped_columns=unmapped_columns,
            suggestions=suggestions
        )
        
        logger.info(f"Mapping completed: {len(mappings)} columns mapped, "
                   f"confidence: {overall_confidence:.2f}")
        
        return result
    
    def _generate_mapping_suggestions(self, mappings: List[ColumnMapping], 
                                    unmapped_columns: List[int], 
                                    headers: List[str]) -> List[str]:
        """Generate suggestions for improving mappings"""
        suggestions = []
        
        # Suggest for unmapped columns
        for col_idx in unmapped_columns:
            if col_idx < len(headers):
                header = headers[col_idx]
                suggestions.append(f"Column {col_idx + 1} '{header}' could not be mapped")
        
        # Suggest improvements for low confidence mappings
        low_confidence_mappings = [m for m in mappings if m.confidence < 0.5]
        if low_confidence_mappings:
            suggestions.append(f"{len(low_confidence_mappings)} columns have low confidence mappings")
        
        # Check for missing required columns
        required_types = self.config.get_required_columns()
        mapped_types = {m.mapped_type for m in mappings}
        missing_required = [col_type for col_type in required_types if col_type not in mapped_types]
        
        if missing_required:
            suggestions.append(f"Missing required columns: {[col.value for col in missing_required]}")
        
        return suggestions


# Convenience function for quick column mapping
def map_columns_quick(headers: List[str]) -> Dict[int, str]:
    """
    Quick column mapping
    
    Args:
        headers: List of header strings
        
    Returns:
        Dictionary mapping column index to column type
    """
    mapper = ColumnMapper()
    mappings = mapper.map_columns_to_types(headers)
    
    result = {}
    for mapping in mappings:
        result[mapping.column_index] = mapping.mapped_type.value
    
    return result 