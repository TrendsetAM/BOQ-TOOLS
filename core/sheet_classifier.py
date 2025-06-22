"""
Sheet Classifier for BOQ Tools
Intelligent sheet classification using heuristic scoring and pattern analysis
"""

import logging
import re
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum
from difflib import SequenceMatcher

from utils.config import get_config, SheetClassification

logger = logging.getLogger(__name__)


class SheetType(Enum):
    """Enumeration for different sheet types"""
    GENERAL_INFO = "general_info"
    SUMMARY = "summary"
    LINE_ITEMS = "line_items"
    REFERENCE = "reference"
    MIXED = "mixed"
    UNKNOWN = "unknown"


@dataclass
class ClassificationResult:
    """Result of sheet classification"""
    sheet_type: SheetType
    confidence: float
    reasoning: List[str]
    scores: Dict[str, float]
    patterns_detected: List[str]
    keyword_matches: List[str]


@dataclass
class PatternInfo:
    """Information about detected patterns"""
    pattern_type: str
    frequency: int
    confidence: float
    description: str


class SheetClassifier:
    """
    Intelligent sheet classifier using heuristic scoring and pattern analysis
    """
    
    def __init__(self):
        """Initialize the sheet classifier"""
        self.config = get_config()
        self._setup_keyword_patterns()
        self._setup_numeric_patterns()
        self._setup_financial_patterns()
        
        logger.info("Sheet Classifier initialized")
    
    def _setup_keyword_patterns(self):
        """Setup keyword patterns for different sheet types"""
        self.keyword_patterns = {
            SheetType.GENERAL_INFO: [
                r"general", r"info", r"information", r"overview", r"project",
                r"details", r"particulars", r"specifications", r"requirements",
                r"preliminary", r"prelim", r"site", r"setup", r"temporary"
            ],
            SheetType.SUMMARY: [
                r"summary", r"total", r"sum", r"grand", r"final", r"overall",
                r"consolidated", r"aggregate", r"breakdown", r"summary of",
                r"total summary", r"project total", r"final total"
            ],
            SheetType.LINE_ITEMS: [
                r"items", r"line", r"boq", r"bill", r"quantities", r"takeoff",
                r"measurement", r"work", r"activity", r"task", r"scope",
                r"detailed", r"breakdown", r"itemized", r"schedule"
            ],
            SheetType.REFERENCE: [
                r"reference", r"ref", r"index", r"table", r"schedule",
                r"specification", r"spec", r"standard", r"code", r"regulation",
                r"appendix", r"annex", r"attachment", r"supporting"
            ]
        }
    
    def _setup_numeric_patterns(self):
        """Setup patterns for numeric content analysis"""
        self.numeric_patterns = {
            "quantity": r"^\d+(\.\d+)?$",
            "currency": r"^\$?\d+(,\d{3})*(\.\d{2})?$",
            "percentage": r"^\d+(\.\d+)?%$",
            "dimension": r"^\d+(\.\d+)?\s*(m|cm|mm|ft|in|kg|ton|l|gal)$",
            "date": r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$",
            "decimal": r"^\d+\.\d+$",
            "integer": r"^\d+$"
        }
    
    def _setup_financial_patterns(self):
        """Setup patterns for financial aggregation detection"""
        self.financial_patterns = {
            "subtotal": [
                r"sub.?total", r"sub total", r"section total", r"group total",
                r"category total", r"sub.?sum", r"partial total"
            ],
            "total": [
                r"^total$", r"^grand total$", r"^final total$", r"^overall total$",
                r"^project total$", r"^sum total$", r"^total amount$"
            ],
            "contingency": [
                r"contingency", r"allowance", r"provisional", r"provisional sum",
                r"prime cost", r"provisional cost", r"allowance sum"
            ],
            "tax": [
                r"tax", r"vat", r"gst", r"sales tax", r"value added tax",
                r"goods and services tax", r"taxation"
            ]
        }
    
    def classify_sheet(self, sheet_data: Dict[str, Any], sheet_name: str) -> ClassificationResult:
        """
        Classify a sheet based on its data and name
        
        Args:
            sheet_data: Dictionary containing sheet metadata and content
            sheet_name: Name of the sheet
            
        Returns:
            ClassificationResult with classification details
        """
        logger.debug(f"Classifying sheet: {sheet_name}")
        
        # Extract data from sheet_data
        content = sheet_data.get('content', [])
        headers = sheet_data.get('headers', [])
        metadata = sheet_data.get('metadata', {})
        
        # Calculate individual scores
        keyword_score = self.score_keywords(sheet_name, content, headers)
        numeric_score = self.calculate_numeric_ratio(content, headers)
        pattern_score = self.detect_patterns(content, headers)
        
        # Calculate weighted score
        total_score = (
            keyword_score['score'] * 0.3 +
            numeric_score['ratio'] * 0.4 +
            pattern_score['score'] * 0.3
        )
        
        # Determine sheet type based on scores
        sheet_type = self._determine_sheet_type(
            keyword_score, numeric_score, pattern_score, total_score
        )
        
        # Build reasoning
        reasoning = self._build_reasoning(
            sheet_type, keyword_score, numeric_score, pattern_score, total_score
        )
        
        # Create result
        result = ClassificationResult(
            sheet_type=sheet_type,
            confidence=total_score,
            reasoning=reasoning,
            scores={
                'keyword': keyword_score['score'],
                'numeric': numeric_score['ratio'],
                'pattern': pattern_score['score'],
                'total': total_score
            },
            patterns_detected=pattern_score['patterns'],
            keyword_matches=keyword_score['matches']
        )
        
        logger.debug(f"Sheet '{sheet_name}' classified as {sheet_type.value} "
                    f"with confidence {total_score:.2f}")
        
        return result
    
    def score_keywords(self, sheet_name: str, content: List[List[str]], 
                      headers: List[str]) -> Dict[str, Any]:
        """
        Score sheet based on keyword matching
        
        Args:
            sheet_name: Name of the sheet
            content: Sheet content as list of rows
            headers: Column headers
            
        Returns:
            Dictionary with score and matches
        """
        score = 0.0
        matches = []
        
        # Check sheet name
        name_lower = sheet_name.lower()
        for sheet_type, patterns in self.keyword_patterns.items():
            for pattern in patterns:
                if re.search(pattern, name_lower, re.IGNORECASE):
                    score += 0.4  # Higher weight for name matches
                    matches.append(f"Sheet name matches {sheet_type.value}: '{pattern}'")
        
        # Check headers
        header_text = " ".join(headers).lower()
        for sheet_type, patterns in self.keyword_patterns.items():
            for pattern in patterns:
                if re.search(pattern, header_text, re.IGNORECASE):
                    score += 0.2
                    matches.append(f"Headers match {sheet_type.value}: '{pattern}'")
        
        # Check content (first few rows)
        content_text = " ".join([str(cell) for row in content[:5] for cell in row]).lower()
        for sheet_type, patterns in self.keyword_patterns.items():
            for pattern in patterns:
                if re.search(pattern, content_text, re.IGNORECASE):
                    score += 0.1
                    matches.append(f"Content matches {sheet_type.value}: '{pattern}'")
        
        # Normalize score
        score = min(score, 1.0)
        
        return {
            'score': score,
            'matches': matches
        }
    
    def calculate_numeric_ratio(self, content: List[List[str]], 
                               headers: List[str]) -> Dict[str, Any]:
        """
        Calculate the ratio of numeric content in the sheet
        
        Args:
            content: Sheet content as list of rows
            headers: Column headers
            
        Returns:
            Dictionary with numeric ratio and analysis
        """
        if not content:
            return {'ratio': 0.0, 'analysis': 'No content'}
        
        total_cells = 0
        numeric_cells = 0
        numeric_columns = []
        
        # Analyze headers for numeric indicators
        for i, header in enumerate(headers):
            header_lower = header.lower()
            if any(keyword in header_lower for keyword in 
                   ['qty', 'quantity', 'amount', 'price', 'rate', 'total', 'no', 'number']):
                numeric_columns.append(i)
        
        # Analyze content
        for row in content:
            for i, cell in enumerate(row):
                if not cell or str(cell).strip() == "":
                    continue
                
                total_cells += 1
                cell_str = str(cell).strip()
                
                # Check if cell is numeric
                is_numeric = False
                
                # Check various numeric patterns
                for pattern_name, pattern in self.numeric_patterns.items():
                    if re.match(pattern, cell_str, re.IGNORECASE):
                        is_numeric = True
                        break
                
                # Additional checks for currency and numbers
                if not is_numeric:
                    # Remove common currency symbols and check
                    clean_cell = re.sub(r'[$,€£¥₹]', '', cell_str)
                    if clean_cell.replace('.', '').replace(',', '').isdigit():
                        is_numeric = True
                
                if is_numeric:
                    numeric_cells += 1
        
        ratio = numeric_cells / max(1, total_cells)
        
        analysis = f"{numeric_cells}/{total_cells} cells are numeric ({ratio:.1%})"
        if numeric_columns:
            analysis += f", {len(numeric_columns)} columns have numeric headers"
        
        return {
            'ratio': ratio,
            'analysis': analysis,
            'numeric_cells': numeric_cells,
            'total_cells': total_cells,
            'numeric_columns': numeric_columns
        }
    
    def detect_patterns(self, content: List[List[str]], 
                       headers: List[str]) -> Dict[str, Any]:
        """
        Detect patterns in sheet content
        
        Args:
            content: Sheet content as list of rows
            headers: Column headers
            
        Returns:
            Dictionary with pattern score and detected patterns
        """
        patterns = []
        score = 0.0
        
        if not content:
            return {'score': 0.0, 'patterns': patterns}
        
        # Detect financial aggregation patterns
        financial_patterns = self._detect_financial_patterns(content, headers)
        patterns.extend(financial_patterns)
        
        # Detect data structure patterns
        structure_patterns = self._detect_structure_patterns(content, headers)
        patterns.extend(structure_patterns)
        
        # Detect repetition patterns
        repetition_patterns = self._detect_repetition_patterns(content)
        patterns.extend(repetition_patterns)
        
        # Calculate pattern score
        if financial_patterns:
            score += 0.4
        if structure_patterns:
            score += 0.3
        if repetition_patterns:
            score += 0.3
        
        return {
            'score': min(score, 1.0),
            'patterns': patterns
        }
    
    def _detect_financial_patterns(self, content: List[List[str]], 
                                  headers: List[str]) -> List[str]:
        """Detect financial aggregation patterns"""
        patterns = []
        
        # Check for financial keywords in headers
        header_text = " ".join(headers).lower()
        for category, keywords in self.financial_patterns.items():
            for keyword in keywords:
                if re.search(keyword, header_text, re.IGNORECASE):
                    patterns.append(f"Financial pattern: {category} in headers")
                    break
        
        # Check for financial patterns in content
        for row in content:
            row_text = " ".join([str(cell) for cell in row]).lower()
            for category, keywords in self.financial_patterns.items():
                for keyword in keywords:
                    if re.search(keyword, row_text, re.IGNORECASE):
                        patterns.append(f"Financial pattern: {category} in content")
                        break
        
        return patterns
    
    def _detect_structure_patterns(self, content: List[List[str]], 
                                  headers: List[str]) -> List[str]:
        """Detect data structure patterns"""
        patterns = []
        
        if not content or not headers:
            return patterns
        
        # Check for consistent column structure
        if len(headers) >= 3:
            patterns.append("Structured data: Multiple columns")
        
        # Check for header row pattern
        if headers and any(header.strip() for header in headers):
            patterns.append("Structured data: Clear headers")
        
        # Check for data consistency
        if len(content) > 1:
            row_lengths = [len(row) for row in content if row]
            if len(set(row_lengths)) <= 2:  # Allow for minor variations
                patterns.append("Structured data: Consistent row structure")
        
        # Check for empty row patterns (section breaks)
        empty_rows = sum(1 for row in content if not any(cell.strip() for cell in row))
        if empty_rows > 0 and empty_rows < len(content) * 0.3:
            patterns.append("Structured data: Section breaks detected")
        
        return patterns
    
    def _detect_repetition_patterns(self, content: List[List[str]]) -> List[str]:
        """Detect repetition patterns in content"""
        patterns = []
        
        if not content:
            return patterns
        
        # Check for repeated values in columns
        for col_idx in range(min(len(row) for row in content if row)):
            column_values = [row[col_idx] for row in content if len(row) > col_idx and row[col_idx].strip()]
            if column_values:
                unique_values = set(column_values)
                repetition_ratio = 1 - (len(unique_values) / len(column_values))
                
                if repetition_ratio > 0.3:
                    patterns.append(f"Repetition pattern: Column {col_idx + 1} has {repetition_ratio:.1%} repetition")
        
        # Check for similar row patterns
        if len(content) > 2:
            similar_rows = 0
            for i in range(len(content) - 1):
                if self._rows_are_similar(content[i], content[i + 1]):
                    similar_rows += 1
            
            if similar_rows > len(content) * 0.2:
                patterns.append("Repetition pattern: Similar row structures detected")
        
        return patterns
    
    def _rows_are_similar(self, row1: List[str], row2: List[str]) -> bool:
        """Check if two rows have similar structure"""
        if not row1 or not row2:
            return False
        
        # Check if both rows have similar data types
        numeric_count1 = sum(1 for cell in row1 if self._is_numeric(cell))
        numeric_count2 = sum(1 for cell in row2 if self._is_numeric(cell))
        
        return abs(numeric_count1 - numeric_count2) <= 1
    
    def _is_numeric(self, cell: str) -> bool:
        """Check if a cell contains numeric data"""
        if not cell or str(cell).strip() == "":
            return False
        
        cell_str = str(cell).strip()
        
        # Check numeric patterns
        for pattern in self.numeric_patterns.values():
            if re.match(pattern, cell_str, re.IGNORECASE):
                return True
        
        # Additional check for currency
        clean_cell = re.sub(r'[$,€£¥₹]', '', cell_str)
        return clean_cell.replace('.', '').replace(',', '').isdigit()
    
    def _determine_sheet_type(self, keyword_score: Dict, numeric_score: Dict, 
                             pattern_score: Dict, total_score: float) -> SheetType:
        """Determine sheet type based on scores"""
        
        # Check for strong keyword matches first
        if keyword_score['score'] > 0.6:
            if any('line_items' in match for match in keyword_score['matches']):
                return SheetType.LINE_ITEMS
            elif any('summary' in match for match in keyword_score['matches']):
                return SheetType.SUMMARY
            elif any('general_info' in match for match in keyword_score['matches']):
                return SheetType.GENERAL_INFO
            elif any('reference' in match for match in keyword_score['matches']):
                return SheetType.REFERENCE
        
        # Check numeric content patterns
        if numeric_score['ratio'] > 0.7:
            if pattern_score['patterns'] and any('financial' in p for p in pattern_score['patterns']):
                return SheetType.SUMMARY
            else:
                return SheetType.LINE_ITEMS
        
        # Check pattern-based classification
        if pattern_score['score'] > 0.6:
            if any('structured' in p for p in pattern_score['patterns']):
                return SheetType.LINE_ITEMS
            elif any('financial' in p for p in pattern_score['patterns']):
                return SheetType.SUMMARY
        
        # Mixed classification for moderate scores
        if total_score > 0.4:
            return SheetType.MIXED
        
        return SheetType.UNKNOWN
    
    def _build_reasoning(self, sheet_type: SheetType, keyword_score: Dict, 
                        numeric_score: Dict, pattern_score: Dict, 
                        total_score: float) -> List[str]:
        """Build reasoning for classification"""
        reasoning = []
        
        reasoning.append(f"Sheet classified as {sheet_type.value} with confidence {total_score:.2f}")
        reasoning.append("")
        
        # Keyword reasoning
        if keyword_score['score'] > 0:
            reasoning.append("Keyword Analysis:")
            reasoning.append(f"  Score: {keyword_score['score']:.2f}")
            for match in keyword_score['matches'][:3]:  # Show first 3 matches
                reasoning.append(f"  - {match}")
        
        # Numeric reasoning
        reasoning.append("")
        reasoning.append("Numeric Content Analysis:")
        reasoning.append(f"  Ratio: {numeric_score['ratio']:.2f}")
        reasoning.append(f"  Analysis: {numeric_score['analysis']}")
        
        # Pattern reasoning
        if pattern_score['patterns']:
            reasoning.append("")
            reasoning.append("Pattern Detection:")
            reasoning.append(f"  Score: {pattern_score['score']:.2f}")
            for pattern in pattern_score['patterns'][:5]:  # Show first 5 patterns
                reasoning.append(f"  - {pattern}")
        
        return reasoning
    
    def get_classification_summary(self, results: List[ClassificationResult]) -> Dict[str, Any]:
        """
        Generate summary of multiple sheet classifications
        
        Args:
            results: List of classification results
            
        Returns:
            Summary dictionary
        """
        summary = {
            'total_sheets': len(results),
            'sheet_types': {},
            'confidence_stats': {
                'average': 0.0,
                'min': 1.0,
                'max': 0.0
            },
            'type_distribution': {}
        }
        
        if not results:
            return summary
        
        # Calculate statistics
        confidences = [r.confidence for r in results]
        summary['confidence_stats']['average'] = sum(confidences) / len(confidences)
        summary['confidence_stats']['min'] = min(confidences)
        summary['confidence_stats']['max'] = max(confidences)
        
        # Count sheet types
        for result in results:
            sheet_type = result.sheet_type.value
            summary['sheet_types'][sheet_type] = summary['sheet_types'].get(sheet_type, 0) + 1
        
        # Calculate distribution
        total = len(results)
        for sheet_type, count in summary['sheet_types'].items():
            summary['type_distribution'][sheet_type] = count / total
        
        return summary


# Convenience function for quick classification
def classify_sheet_quick(sheet_data: Dict[str, Any], sheet_name: str) -> Tuple[str, float]:
    """
    Quick sheet classification
    
    Args:
        sheet_data: Sheet data dictionary
        sheet_name: Name of the sheet
        
    Returns:
        Tuple of (sheet_type, confidence)
    """
    classifier = SheetClassifier()
    result = classifier.classify_sheet(sheet_data, sheet_name)
    return result.sheet_type.value, result.confidence 