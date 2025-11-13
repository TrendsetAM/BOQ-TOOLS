"""
Column Mapper for BOQ Tools
Intelligent column mapping with header detection and confidence scoring
"""

import logging
import re
import difflib
import json
import os
from typing import Dict, List, Tuple, Optional, Any, Set
from dataclasses import dataclass
from enum import Enum

from utils.config import get_config, ColumnType

try:
    from appdirs import user_config_dir
except ImportError:
    user_config_dir = None

logger = logging.getLogger(__name__)


class HeaderDetectionMethod(Enum):
    """Methods for header row detection"""
    KEYWORD_MATCH = "keyword_match"
    DATA_TYPE_PATTERN = "data_type_pattern"
    POSITIONAL_LOGIC = "positional_logic"


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
    
    # Default canonical required type mapping
    DEFAULT_CANONICAL_HEADER_MAP = {
        'description': ["description", "desc", "item description", "work description", "item", "activity", "task"],
        'quantity': ["quantity", "qty", "quant", "qty.", "number", "count", "nos"],
        'unit_price': ["unit price", "unit_price", "price/unit", "unitprice", "rate", "unit rate", "price per unit"],
        'total_price': ["total price", "total_price", "amount", "price", "total", "sum", "cost", "value"],
        'unit': ["unit", "uom", "measure", "unit of measure", "units", "measurement"],
        'code': ["code", "item code", "ref code", "reference code", "item no", "item number", "boq code", "schedule no"],
        'scope': ["scope"],
        'manhours': ["manhours", "man hours", "labour ore/u.m.", "labor ore/u.m."],
        'wage': ["euro/hour", "wage", "hourly rate", "labour euro/hour", "labor euro/hour"]
    }
    
    def __init__(self, max_header_rows: int = 20):
        """
        Initialize the column mapper
        
        Args:
            max_header_rows: Maximum number of rows to check for headers (default: 20)
        """
        self.config = get_config()
        self.max_header_rows = max_header_rows
        self._setup_patterns()
        self._load_canonical_mappings()
        
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
    
    @staticmethod
    def get_user_config_path(filename: str) -> str:
        """Get a user-writable config path for the given filename."""
        app_name = "BOQ-TOOLS"
        if user_config_dir:
            config_dir = user_config_dir(app_name)
        else:
            # Fallback: use platform-specific logic
            if os.name == 'nt':
                config_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), app_name)
            else:
                config_dir = os.path.join(os.path.expanduser('~/.config'), app_name)
        os.makedirs(config_dir, exist_ok=True)
        return os.path.join(config_dir, filename)

    def _load_canonical_mappings(self):
        """Load canonical mappings from user config or use defaults. On first run, copy default to user config."""
        # Path to user config file
        self.canonical_mappings_file = self.get_user_config_path('canonical_mappings.json')
        default_path = os.path.join(os.path.dirname(__file__), '..', 'config', 'canonical_mappings.json')
        try:
            if not os.path.exists(self.canonical_mappings_file):
                # Copy default file to user config dir
                if os.path.exists(default_path):
                    with open(default_path, 'r', encoding='utf-8') as fsrc:
                        default_data = fsrc.read()
                    with open(self.canonical_mappings_file, 'w', encoding='utf-8') as fdst:
                        fdst.write(default_data)
                    logger.info(f"Copied default canonical mappings to {self.canonical_mappings_file}")
                else:
                    # If default file missing, use hardcoded defaults
                    with open(self.canonical_mappings_file, 'w', encoding='utf-8') as fdst:
                        json.dump(self.DEFAULT_CANONICAL_HEADER_MAP, fdst, indent=2, ensure_ascii=False)
                    logger.info(f"Created new canonical mappings file with defaults at {self.canonical_mappings_file}")
            # Load from user config
            with open(self.canonical_mappings_file, 'r', encoding='utf-8') as f:
                self.CANONICAL_HEADER_MAP = json.load(f)
            logger.info(f"Loaded canonical mappings from {self.canonical_mappings_file}")
        except Exception as e:
            logger.warning(f"Failed to load canonical mappings: {e}, using defaults")
            self.CANONICAL_HEADER_MAP = self.DEFAULT_CANONICAL_HEADER_MAP.copy()
        # Create lookup dictionary
        self.CANONICAL_TYPE_LOOKUP = {v: k for k, vals in self.CANONICAL_HEADER_MAP.items() for v in vals}
    
    def _save_canonical_mappings(self):
        """Save canonical mappings to user config file"""
        try:
            os.makedirs(os.path.dirname(self.canonical_mappings_file), exist_ok=True)
            with open(self.canonical_mappings_file, 'w', encoding='utf-8') as f:
                json.dump(self.CANONICAL_HEADER_MAP, f, indent=2, ensure_ascii=False)
            logger.info(f"Saved canonical mappings to {self.canonical_mappings_file}")
        except Exception as e:
            logger.error(f"Failed to save canonical mappings: {e}")
    
    def update_canonical_mapping(self, original_header: str, mapped_type: str):
        """
        Update canonical mappings when user confirms a column mapping
        
        Args:
            original_header: The original header text
            mapped_type: The confirmed mapped type (e.g., 'description', 'quantity', etc.)
        """
        if mapped_type not in self.CANONICAL_HEADER_MAP:
            logger.warning(f"Unknown mapped type: {mapped_type}")
            return
        
        # Normalize the header for storage
        normalized_header = original_header.strip()
        
        # Add to canonical mappings if not already present
        if normalized_header not in self.CANONICAL_HEADER_MAP[mapped_type]:
            self.CANONICAL_HEADER_MAP[mapped_type].append(normalized_header)
            self.CANONICAL_TYPE_LOOKUP[normalized_header] = mapped_type
            self._save_canonical_mappings()
            logger.info(f"Added '{normalized_header}' to canonical mappings for '{mapped_type}'")
    
    def get_canonical_mappings(self) -> Dict[str, List[str]]:
        """Get current canonical mappings"""
        return self.CANONICAL_HEADER_MAP.copy()
    

    
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
        search_rows = min(self.max_header_rows, len(sheet_data))
        debug_rows = []  # Collect debug info for each candidate
        keyword_candidates = []  # Track keyword-based candidates separately
        
        for row_index in range(search_rows):
            row = sheet_data[row_index]
            if not row:
                continue
            methods = [
                self._detect_by_keywords,
                self._detect_by_data_patterns,
                self._detect_by_positional_logic
            ]
            for method in methods:
                try:
                    result = method(row, row_index, sheet_data)
                    if result:
                        debug_rows.append({
                            'row_index': row_index,
                            'method': method.__name__,
                            'confidence': result.confidence,
                            'reasoning': result.reasoning,
                            'headers': row
                        })
                        
                        # Track keyword-based candidates separately for tie-breaking
                        if method == self._detect_by_keywords:
                            keyword_candidates.append(result)
                        
                        # Simple scoring: highest confidence wins
                        if result.confidence > best_score:
                            best_score = result.confidence
                            best_header = result
                            
                except Exception as e:
                    logger.warning(f"Error in header detection method {method.__name__}: {e}")
        
        # Additional tie-breaker: if we have multiple keyword candidates with similar scores,
        # prefer the one with more keyword matches
        if len(keyword_candidates) > 1:
            # Find the keyword candidate with the highest score
            best_keyword = max(keyword_candidates, key=lambda x: x.confidence)
            if (best_keyword.confidence >= best_score - 0.05 and  # Very close to best score
                best_header and 
                best_header.method != HeaderDetectionMethod.KEYWORD_MATCH):
                # Prefer the keyword match if it's very close to the best score
                best_score = best_keyword.confidence
                best_header = best_keyword
        
        # Log all candidate rows and their scores
        if debug_rows:
            logger.debug("Header row candidates:")
            for info in debug_rows:
                logger.debug(f"Row {info['row_index']} ({info['method']}): confidence={info['confidence']:.2f}, headers={info['headers']}, reasoning={info['reasoning']}")
        
        if not best_header:
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
        
        # After finding the best header row, check for hierarchical headers
        if best_header:
            # Check if there are parent headers above it
            if best_header.row_index > 0:
                enhanced_header = self._enhance_header_with_parent_row(best_header, sheet_data)
                if enhanced_header:
                    best_header = enhanced_header
            
            # Check if the header row itself has merged cells and subheaders below it
            enhanced_header = self._enhance_header_with_subheader_row(best_header, sheet_data)
            if enhanced_header:
                best_header = enhanced_header
        
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
    
    def _enhance_header_with_parent_row(self, header_info: HeaderRowInfo, sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """
        Check if there's a parent row above the header row that contains merged cells
        Enhanced with logic from former Method 4 for better hierarchical detection
        Only applies enhancement if current headers are not already well-recognized
        
        Args:
            header_info: The detected header row info
            sheet_data: Complete sheet data
            
        Returns:
            Enhanced HeaderRowInfo if parent row with merged cells detected, None otherwise
        """
        row_index = header_info.row_index
        current_row = sheet_data[row_index]  # This is the detected header row (subheaders)
        
        # Check if there's a previous row to use as parent headers
        if row_index == 0:
            return None
        

        
        parent_row = sheet_data[row_index - 1]  # This could be the parent header row
        
        # Enhanced detection logic (migrated from Method 4)
        enhancement_confidence = 0.0
        enhancement_reasoning = []
        
        # Check for merged cell indicators (empty cells between content)
        empty_cells = sum(1 for cell in parent_row if not cell or not str(cell).strip())
        total_cells = len(parent_row)
        
        if empty_cells > 0 and empty_cells < total_cells:
            enhancement_confidence = 0.2
            enhancement_reasoning.append(f"Potential merged cells: {empty_cells}/{total_cells} empty cells")
        
        # Check if parent row has some content (indicating it's a header row)
        parent_content = sum(1 for cell in parent_row if cell and str(cell).strip())
        
        if parent_content == 0:
            # No content in parent row
            return None
        
        # Enhanced parent-subheader relationship detection (from Method 4)
        parent_non_empty = [i for i, cell in enumerate(parent_row) if cell and str(cell).strip()]
        current_non_empty = [i for i, cell in enumerate(current_row) if cell and str(cell).strip()]
        
        # Look for patterns where parent row has fewer cells that might span multiple columns
        if len(parent_non_empty) > 0 and len(current_non_empty) > len(parent_non_empty):
            # Check if parent row cells could be parent headers
            for i, cell in enumerate(parent_row):
                if cell and str(cell).strip():
                    cell_lower = str(cell).strip().lower()
                    # Look for common parent header terms
                    if any(term in cell_lower for term in ['labour', 'labor', 'material', 'equipment', 'cost', 'price', 'analysis']):
                        enhancement_confidence = max(enhancement_confidence, 0.4)
                        enhancement_reasoning.append(f"Potential parent header '{cell}' with subheaders below")
            
            # Enhanced detection for labor-related hierarchical headers (from Method 4)
            combined_headers = self._create_hierarchical_headers(parent_row, current_row)
            labor_matches = 0
            for combined_header in combined_headers:
                if combined_header:
                    combined_lower = combined_header.lower()
                    if any(term in combined_lower for term in ['ore/u.m.', 'euro/hour', 'manhours', 'wage']):
                        labor_matches += 1
            
            if labor_matches > 0:
                # Boost confidence significantly for labor hierarchical headers
                enhancement_confidence = max(enhancement_confidence, 0.9)
                enhancement_reasoning.append(f"Found {labor_matches} labor-related hierarchical headers")
                
                # Use the enhanced headers
                enhanced_headers = combined_headers
            else:
                # Use selective hierarchical headers
                enhanced_headers = self._create_selective_hierarchical_headers(parent_row, current_row)
        else:
            # Use selective hierarchical headers
            enhanced_headers = self._create_selective_hierarchical_headers(parent_row, current_row)
        
        # Check if we actually created any hierarchical headers
        has_hierarchical = any(' ' in header for header in enhanced_headers if header)
        
        if has_hierarchical or enhancement_confidence >= 0.25:
            # Boost the original confidence with enhancement confidence
            boosted_confidence = min(1.0, header_info.confidence + (enhancement_confidence * 0.3))
            
            # Add reasoning about the enhancement
            enhanced_reasoning = header_info.reasoning.copy()
            enhanced_reasoning.extend(enhancement_reasoning)
            enhanced_reasoning.append("Enhanced with hierarchical headers from parent row merged cells")
            
            return HeaderRowInfo(
                row_index=header_info.row_index,
                confidence=boosted_confidence,
                method=header_info.method,
                reasoning=enhanced_reasoning,
                headers=enhanced_headers,
                is_merged=True
            )
        
        return None
    
    def _enhance_header_with_subheader_row(self, header_info: HeaderRowInfo, sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """
        Check if the header row itself has merged cells and subheaders below it
        Enhanced with logic from former Method 4 for better hierarchical detection
        Only applies enhancement if current headers are not already well-recognized
        
        Args:
            header_info: The detected header row info
            sheet_data: Complete sheet data
            
        Returns:
            Enhanced HeaderRowInfo if subheader row with merged cells detected, None otherwise
        """
        row_index = header_info.row_index
        current_row = sheet_data[row_index]  # This is the detected header row (parent headers)
        
        # Check if there's a next row to use as subheaders
        if row_index >= len(sheet_data) - 1:
            return None
        

        
        subheader_row = sheet_data[row_index + 1]  # This could be the subheader row
        
        # Enhanced detection logic (migrated from Method 4)
        enhancement_confidence = 0.0
        enhancement_reasoning = []
        
        # Check for merged cell indicators (empty cells between content)
        empty_cells = sum(1 for cell in current_row if not cell or not str(cell).strip())
        total_cells = len(current_row)
        
        if empty_cells > 0 and empty_cells < total_cells:
            enhancement_confidence = 0.2
            enhancement_reasoning.append(f"Potential merged cells: {empty_cells}/{total_cells} empty cells")
        
        # Check if current row has some content (indicating it's a header row)
        current_content = sum(1 for cell in current_row if cell and str(cell).strip())
        
        if current_content == 0:
            # No content in current row
            return None
        
        # Check if subheader row has content
        subheader_content = sum(1 for cell in subheader_row if cell and str(cell).strip())
        
        if subheader_content == 0:
            # No content in subheader row
            return None
        
        # Enhanced parent-subheader relationship detection (from Method 4)
        current_non_empty = [i for i, cell in enumerate(current_row) if cell and str(cell).strip()]
        next_non_empty = [i for i, cell in enumerate(subheader_row) if cell and str(cell).strip()]
        
        # Look for patterns where current row has fewer cells that might span multiple columns
        if len(current_non_empty) > 0 and len(next_non_empty) > len(current_non_empty):
            # Check if current row cells could be parent headers
            for i, cell in enumerate(current_row):
                if cell and str(cell).strip():
                    cell_lower = str(cell).strip().lower()
                    # Look for common parent header terms
                    if any(term in cell_lower for term in ['labour', 'labor', 'material', 'equipment', 'cost', 'price', 'analysis']):
                        enhancement_confidence = max(enhancement_confidence, 0.4)
                        enhancement_reasoning.append(f"Potential parent header '{cell}' with subheaders below")
            
            # Enhanced detection for labor-related hierarchical headers (from Method 4)
            combined_headers = self._create_hierarchical_headers(current_row, subheader_row)
            labor_matches = 0
            for combined_header in combined_headers:
                if combined_header:
                    combined_lower = combined_header.lower()
                    if any(term in combined_lower for term in ['ore/u.m.', 'euro/hour', 'manhours', 'wage']):
                        labor_matches += 1
            
            if labor_matches > 0:
                # Boost confidence significantly for labor hierarchical headers
                enhancement_confidence = max(enhancement_confidence, 0.9)
                enhancement_reasoning.append(f"Found {labor_matches} labor-related hierarchical headers")
                
                # Use the enhanced headers
                enhanced_headers = combined_headers
            else:
                # Use selective hierarchical headers
                enhanced_headers = self._create_selective_hierarchical_headers(current_row, subheader_row)
        else:
            # Use selective hierarchical headers
            enhanced_headers = self._create_selective_hierarchical_headers(current_row, subheader_row)
        
        # Check if we actually created any hierarchical headers
        has_hierarchical = any(' ' in header for header in enhanced_headers if header)
        
        if has_hierarchical or enhancement_confidence >= 0.25:
            # Boost the original confidence with enhancement confidence
            boosted_confidence = min(1.0, header_info.confidence + (enhancement_confidence * 0.3))
            
            # Add reasoning about the enhancement
            enhanced_reasoning = header_info.reasoning.copy()
            enhanced_reasoning.extend(enhancement_reasoning)
            enhanced_reasoning.append("Enhanced with hierarchical headers from subheader row merged cells")
            
            return HeaderRowInfo(
                row_index=header_info.row_index,  # Keep the same row index
                confidence=boosted_confidence,
                method=header_info.method,
                reasoning=enhanced_reasoning,
                headers=enhanced_headers,
                is_merged=True
            )
        
        return None
    
    def _create_selective_hierarchical_headers(self, parent_row: List[str], subheader_row: List[str]) -> List[str]:
        """
        Create hierarchical headers only for merged cell sections, keep original headers for others
        Enhanced to preserve individual well-recognized headers even during enhancement
        
        Args:
            parent_row: Parent header row (may have empty cells indicating merged cells)
            subheader_row: Subheader row
            
        Returns:
            List of headers with hierarchical structure only where needed
        """
        enhanced_headers = []
        
        # Ensure both rows have the same length
        max_length = max(len(parent_row), len(subheader_row))
        parent_row = parent_row + [""] * (max_length - len(parent_row))
        subheader_row = subheader_row + [""] * (max_length - len(subheader_row))
        
        # Identify parent header spans
        parent_spans = self._identify_parent_spans(parent_row, subheader_row)
        
        for i, subheader in enumerate(subheader_row):
            # Check if the subheader is already a well-recognized keyword
            subheader_recognized = False
            if subheader and str(subheader).strip():
                subheader_lower = str(subheader).strip().lower()
                for column_type, keywords in self.CANONICAL_HEADER_MAP.items():
                    for keyword in keywords:
                        if keyword.lower() == subheader_lower:
                            subheader_recognized = True
                            break
                    if subheader_recognized:
                        break
            
            # Check if the parent header is already a well-recognized keyword
            parent_recognized = False
            if i < len(parent_row) and parent_row[i] and str(parent_row[i]).strip():
                parent_lower = str(parent_row[i]).strip().lower()
                for column_type, keywords in self.CANONICAL_HEADER_MAP.items():
                    for keyword in keywords:
                        if keyword.lower() == parent_lower:
                            parent_recognized = True
                            break
                    if parent_recognized:
                        break
            
            # Check if this column is part of a merged cell span
            parent_header = ""
            is_in_span = False
            
            for span_start, span_end, span_header in parent_spans:
                if span_start <= i <= span_end and span_end > span_start:  # Only for multi-column spans
                    parent_header = span_header
                    is_in_span = True
                    break
            
            # SMART HEADER SELECTION: Preserve recognized headers
            if subheader_recognized:
                # Subheader is recognized - use it as-is, don't enhance
                enhanced_header = str(subheader).strip()
            elif parent_recognized and not is_in_span:
                # Parent header is recognized and not in merged span - use it as-is
                enhanced_header = str(parent_row[i]).strip()
            elif is_in_span and subheader and str(subheader).strip():
                # This is part of a merged cell span and subheader is not recognized, create hierarchical header
                subheader_str = str(subheader).strip()
                enhanced_header = f"{parent_header} {subheader_str}"
            elif not is_in_span and parent_row[i] and str(parent_row[i]).strip():
                # This is not in a merged span, use the original parent header
                enhanced_header = str(parent_row[i]).strip()
            elif subheader and str(subheader).strip():
                # Fallback to subheader if available
                enhanced_header = str(subheader).strip()
            elif parent_header:
                # No subheader but has parent header
                enhanced_header = parent_header
            else:
                # Empty
                enhanced_header = ""
            
            enhanced_headers.append(enhanced_header)
        
        return enhanced_headers
    
    def _detect_by_keywords(self, row: List[str], row_index: int, 
                           sheet_data: List[List[str]]) -> Optional[HeaderRowInfo]:
        """Detect header row by exact keyword matching using canonical mappings dictionary"""
        if not row:
            return None
        
        score = 0.0
        matches = []
        keyword_count = 0
        
        # Check each cell for header keywords
        for cell in row:
            if not cell:
                continue
            
            cell_lower = str(cell).lower().strip()
            
            # Check against canonical mappings dictionary (learned keywords)
            for column_type, keywords in self.CANONICAL_HEADER_MAP.items():
                for keyword in keywords:
                    # Changed from substring matching to exact matching
                    if keyword.lower() == cell_lower:
                        # Get weight from config for this column type
                        col_type_enum = None
                        for ct in self.config.get_all_column_types():
                            if ct.value == column_type:
                                col_type_enum = ct
                                break
                        
                        if col_type_enum:
                            mapping = self.config.get_column_mapping(col_type_enum)
                            weight = mapping.weight if mapping else 1.0
                        else:
                            weight = 1.0  # Default weight for unknown types
                        
                        score += weight
                        keyword_count += 1
                        matches.append(f"'{cell}' exactly matches {column_type}")
        
        # Improved scoring algorithm that prioritizes multiple keyword matches
        if keyword_count > 0:
            # Base score from keyword matches (not normalized by row length)
            base_score = min(score / 10.0, 0.6)  # Cap base score at 0.6
            
            # Bonus for multiple keyword matches
            keyword_bonus = min(keyword_count * 0.15, 0.4)  # Up to 0.4 bonus for multiple keywords
            
            # Penalty for too many empty cells (but not as severe)
            non_empty_cells = sum(1 for cell in row if cell and str(cell).strip())
            empty_penalty = max(0, (len(row) - non_empty_cells) / len(row) * 0.2)  # Max 0.2 penalty
            
            score = base_score + keyword_bonus - empty_penalty
            
            # Ensure score is within bounds
            score = max(0.0, min(score, 1.0))
        
        if score > 0.3:  # Threshold for keyword detection
            reasoning = [f"Exact keyword matches: {', '.join(matches[:3])}"]
            if keyword_count > 3:
                reasoning.append(f"Multiple exact keyword matches: {keyword_count} keywords found")
            
            return HeaderRowInfo(
                row_index=row_index,
                confidence=score,
                method=HeaderDetectionMethod.KEYWORD_MATCH,
                reasoning=reasoning,
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
    

    
    def _create_hierarchical_headers(self, parent_row: List[str], subheader_row: List[str]) -> List[str]:
        """
        Create combined headers from parent and subheader rows, simulating merged cell behavior
        
        Args:
            parent_row: Parent header row
            subheader_row: Subheader row
            
        Returns:
            List of combined header strings
        """
        combined_headers = []
        
        # Ensure both rows have the same length
        max_length = max(len(parent_row), len(subheader_row))
        parent_row = parent_row + [""] * (max_length - len(parent_row))
        subheader_row = subheader_row + [""] * (max_length - len(subheader_row))
        
        # Step 1: Identify parent header spans (simulate merged cells)
        parent_spans = self._identify_parent_spans(parent_row, subheader_row)
        
        # Step 2: Create combined headers based on spans
        for i, subheader in enumerate(subheader_row):
            # Find which parent header spans this column
            parent_header = ""
            for span_start, span_end, span_header in parent_spans:
                if span_start <= i <= span_end:
                    parent_header = span_header
                    break
            
            # Create combined header
            if subheader and str(subheader).strip():
                subheader_str = str(subheader).strip()
                if parent_header:
                    combined_header = f"{parent_header} {subheader_str}"
                else:
                    combined_header = subheader_str
            else:
                combined_header = parent_header if parent_header else ""
            
            combined_headers.append(combined_header)
        
        return combined_headers
    
    def _identify_parent_spans(self, parent_row: List[str], subheader_row: List[str]) -> List[Tuple[int, int, str]]:
        """
        Identify the spans of parent headers, simulating merged cell behavior
        
        Args:
            parent_row: Parent header row
            subheader_row: Subheader row
            
        Returns:
            List of tuples (start_col, end_col, header_text) for each parent header span
        """
        spans = []
        i = 0
        
        while i < len(parent_row):
            parent_cell = parent_row[i]
            
            if parent_cell and str(parent_cell).strip():
                parent_header = str(parent_cell).strip()
                span_start = i
                
                # Find the end of this parent header's span
                # A parent header spans until:
                # 1. We find another non-empty parent header, OR
                # 2. We find empty subheaders (indicating end of logical group), OR
                # 3. We reach the end of the row
                
                span_end = i
                j = i + 1
                consecutive_empty_subheaders = 0
                
                while j < len(parent_row):
                    # If we find another parent header, the current span ends
                    if parent_row[j] and str(parent_row[j]).strip():
                        break
                    
                    # If we find a subheader, extend the span
                    if j < len(subheader_row) and subheader_row[j] and str(subheader_row[j]).strip():
                        span_end = j
                        consecutive_empty_subheaders = 0
                    else:
                        consecutive_empty_subheaders += 1
                        # If we have too many consecutive empty subheaders, stop the span
                        if consecutive_empty_subheaders >= 2:
                            break
                    
                    j += 1
                
                spans.append((span_start, span_end, parent_header))
                i = span_end + 1
            else:
                i += 1
        
        return spans
    
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
    
    def _normalize_header(self, header):
        return re.sub(r'[^a-z0-9]', '', header.strip().lower())

    def _canonical_type_for_header(self, header):
        # Use exact matching instead of normalized matching
        header_lower = header.strip().lower()
        for canonical, variants in self.CANONICAL_HEADER_MAP.items():
            for variant in variants:
                if header_lower == variant.lower():
                    return canonical
        # Remove fuzzy match fallback to enforce exact matching only
        return None

    def map_columns_to_types(self, headers: List[str]) -> List[ColumnMapping]:
        """
        Map columns to BOQ types using only canonical mapping - no fuzzy matching
        """
        logger.debug(f"Mapping {len(headers)} columns to BOQ types (canonical only)")
        all_mappings = []
        
        for col_idx, original_header in enumerate(headers):
            if not original_header or not str(original_header).strip():
                # Create mapping for empty headers as IGNORE to maintain column index consistency
                mapping = ColumnMapping(
                    column_index=col_idx,
                    original_header=original_header or "",
                    normalized_header="",
                    mapped_type=ColumnType.IGNORE,
                    confidence=0.0,
                    alternatives=[(ColumnType.IGNORE, 0.0)],
                    reasoning=["Empty header - mapped to ignore"]
                )
                all_mappings.append(mapping)
                continue
                
            # Try canonical mapping only
            canonical_type = self._canonical_type_for_header(original_header)
            if canonical_type:
                # 100% confidence for canonical match
                col_type = getattr(ColumnType, canonical_type.upper(), ColumnType.IGNORE)
                mapping = ColumnMapping(
                    column_index=col_idx,
                    original_header=original_header,
                    normalized_header=self._normalize_header(original_header),
                    mapped_type=col_type,
                    confidence=1.0,
                    alternatives=[(col_type, 1.0)],
                    reasoning=[f"Canonical match for '{original_header}' as '{col_type.value}'"]
                )
                all_mappings.append(mapping)
            else:
                # No canonical match - map to IGNORE for manual user correction
                mapping = ColumnMapping(
                    column_index=col_idx,
                    original_header=original_header,
                    normalized_header=self._normalize_header(original_header),
                    mapped_type=ColumnType.IGNORE,
                    confidence=0.0,
                    alternatives=[(ColumnType.IGNORE, 0.0)],
                    reasoning=[f"No canonical match found for '{original_header}' - requires manual mapping"]
                )
                all_mappings.append(mapping)
        
        # Enforce uniqueness for required columns (keep only the first match for each required type)
        required_types = {ColumnType.DESCRIPTION, ColumnType.QUANTITY, ColumnType.UNIT_PRICE, 
                          ColumnType.TOTAL_PRICE, ColumnType.UNIT, ColumnType.CODE}
        seen_required = set()
        
        for mapping in all_mappings:
            if mapping.mapped_type in required_types:
                if mapping.mapped_type in seen_required:
                    # Demote duplicate required column to IGNORE
                    original_type = mapping.mapped_type.value
                    mapping.mapped_type = ColumnType.IGNORE
                    mapping.confidence = 0.0
                    mapping.reasoning.append(f"Demoted from '{original_type}' to 'ignore' - duplicate required column")
                    logger.debug(f"Demoted column '{mapping.original_header}' from {original_type} to ignore (duplicate required column)")
                else:
                    seen_required.add(mapping.mapped_type)
        
        return all_mappings
    

    

    
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
        Get alternative mappings for columns - simplified to only return all column types
        
        Args:
            headers: List of header strings
            
        Returns:
            Dictionary mapping column index to all possible column types
        """
        alternatives = {}
        
        for col_idx, original_header in enumerate(headers):
            if not original_header or not str(original_header).strip():
                # Include empty headers in alternatives as IGNORE only
                alternatives[col_idx] = [(ColumnType.IGNORE, 0.0)]
                continue
            
            # Check if there's a canonical match
            canonical_type = self._canonical_type_for_header(original_header)
            
            if canonical_type:
                # If there's a canonical match, only return that one
                col_type = getattr(ColumnType, canonical_type.upper(), ColumnType.IGNORE)
                alternatives[col_idx] = [(col_type, 1.0)]
            else:
                # No canonical match - return all column types for manual selection
                column_alternatives = []
                for col_type in self.config.get_all_column_types():
                    column_alternatives.append((col_type, 0.0))  # Equal weight for manual selection
                
                alternatives[col_idx] = column_alternatives
        
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
        
        # CONSISTENCY FIX: Use the same enhancement logic as manual selection
        # This ensures auto-detection and manual selection produce identical headers
        logger.info(f"Applying consistent enhancement logic to auto-detected header row {header_info.row_index}")
        
        # Use the forced header processing method to get consistent enhancement
        return self.process_sheet_mapping_with_forced_header(sheet_data, header_info.row_index)
    
    def process_sheet_mapping_with_forced_header(self, sheet_data: List[List[str]], header_row_index: int) -> MappingResult:
        """
        Complete sheet mapping process with user-specified header row
        
        Args:
            sheet_data: Sheet data as list of rows
            header_row_index: 0-based index of the header row (user-specified)
            
        Returns:
            MappingResult with complete mapping information
        """
        logger.info(f"Processing sheet mapping with forced header row {header_row_index} for {len(sheet_data)} rows")
        
        # Validate header row index
        if header_row_index < 0 or header_row_index >= len(sheet_data):
            logger.error(f"Invalid header row index {header_row_index} for sheet with {len(sheet_data)} rows")
            return MappingResult(
                header_row=HeaderRowInfo(
                    row_index=header_row_index,
                    confidence=0.0,
                    method=HeaderDetectionMethod.KEYWORD_MATCH,
                    reasoning=[f"Invalid header row index {header_row_index}"],
                    headers=[],
                    is_merged=False
                ),
                mappings=[],
                overall_confidence=0.0,
                unmapped_columns=[],
                suggestions=[f"Invalid header row index {header_row_index}"]
            )
        
        # Get headers from specified row
        headers = sheet_data[header_row_index]
        
        # Create header info with manual method
        header_info = HeaderRowInfo(
            row_index=header_row_index,
            confidence=1.0,  # User specified = 100% confidence
            method=HeaderDetectionMethod.KEYWORD_MATCH,  # Use existing enum value
            reasoning=["User manually specified header row"],
            headers=headers,
            is_merged=False
        )
        
        if not headers or not any(str(h).strip() for h in headers):
            logger.warning(f"Header row {header_row_index} appears to be empty or invalid")
            return MappingResult(
                header_row=header_info,
                mappings=[],
                overall_confidence=0.0,
                unmapped_columns=list(range(len(headers))),
                suggestions=["Selected header row appears to be empty"]
            )
        
        # Apply sub-header processing similar to find_header_row method
        # Check if there are parent headers above it
        if header_info.row_index > 0:
            enhanced_header = self._enhance_header_with_parent_row(header_info, sheet_data)
            if enhanced_header:
                header_info = enhanced_header
        
        # Check if the header row itself has merged cells and subheaders below it
        enhanced_header = self._enhance_header_with_subheader_row(header_info, sheet_data)
        if enhanced_header:
            header_info = enhanced_header
        
        # Map columns to types using the potentially enhanced headers
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
        
        logger.info(f"Forced header mapping completed: {len(mappings)} columns mapped, "
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