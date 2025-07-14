"""
Data Validator for BOQ Tools
Comprehensive validation with mathematical consistency, data type validation, business rule validation, and confidence scoring.
"""

import logging
import re
from typing import Dict, List, Tuple, Optional, Any, Set
from dataclasses import dataclass
from enum import Enum
from decimal import Decimal, InvalidOperation

from utils.config import get_config, ColumnType

logger = logging.getLogger(__name__)


class ValidationLevel(Enum):
    """Validation severity levels"""
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class ValidationType(Enum):
    """Types of validation checks"""
    MATHEMATICAL = "mathematical"
    DATA_TYPE = "data_type"
    BUSINESS_RULE = "business_rule"
    CONSISTENCY = "consistency"


@dataclass
class ValidationIssue:
    """Individual validation issue"""
    row_index: int
    column_index: Optional[int]
    validation_type: ValidationType
    level: ValidationLevel
    message: str
    expected_value: Optional[Any]
    actual_value: Optional[Any]
    suggestion: Optional[str]


@dataclass
class ValidationResult:
    """Result of validation process"""
    issues: List[ValidationIssue]
    summary: Dict[ValidationLevel, int]
    overall_score: float
    confidence_factors: Dict[str, float]
    suggestions: List[str]


@dataclass
class MathematicalCheck:
    """Mathematical consistency check result"""
    row_index: int
    quantity: Optional[float]
    unit_price: Optional[float]
    total_price: Optional[float]
    calculated_total: Optional[float]
    difference: Optional[float]
    tolerance: float
    is_valid: bool


class DataValidator:
    """
    Comprehensive data validator with mathematical consistency and business rules
    """
    
    def __init__(self, tolerance_percentage: float = 0.01):
        """
        Initialize the data validator
        
        Args:
            tolerance_percentage: Tolerance for mathematical consistency (default 1%)
        """
        self.config = get_config()
        self.tolerance_percentage = tolerance_percentage
        self._setup_patterns()
        
        logger.info("Data Validator initialized")
    
    def _setup_patterns(self):
        """Setup patterns for validation"""
        # Currency patterns
        self.currency_patterns = [
            r'^\$[\d,]+\.?\d*$',  # $1,234.56
            r'^€[\d,]+\.?\d*$',   # €1,234.56
            r'^£[\d,]+\.?\d*$',   # £1,234.56
            r'^¥[\d,]+\.?\d*$',   # ¥1,234.56
            r'^₹[\d,]+\.?\d*$',   # ₹1,234.56
        ]
        
        # Number patterns
        self.number_patterns = [
            r'^[\d,]+\.?\d*$',    # 1,234.56
            r'^[\d]+\.?\d*$',     # 1234.56
            r'^[\d,]+\.?\d*%$',   # 1,234.56%
        ]
        
        # Unit patterns
        self.unit_patterns = [
            r'^m[²³]$',           # m², m³
            r'^sq\.?m$',          # sq.m, sqm
            r'^cu\.?m$',          # cu.m, cum
            r'^kg$',              # kg
            r'^ton$',             # ton
            r'^l$',               # l
            r'^gal$',             # gal
            r'^pcs$',             # pcs
            r'^nos$',             # nos
            r'^units$',           # units
        ]
    
    def validate_mathematical_consistency(self, sheet_data: List[List[str]], 
                                        column_mapping: Dict[int, ColumnType]) -> List[MathematicalCheck]:
        """
        Validate mathematical consistency of calculations
        
        Args:
            sheet_data: Sheet data as list of rows
            column_mapping: Dictionary mapping column index to ColumnType
            
        Returns:
            List of mathematical check results
        """
        checks = []
        
        # Find relevant columns
        quantity_col: Optional[int] = None
        unit_price_col: Optional[int] = None
        total_price_col: Optional[int] = None
        
        for col_idx, col_type in column_mapping.items():
            if col_type == ColumnType.QUANTITY:
                quantity_col = col_idx
            elif col_type == ColumnType.UNIT_PRICE:
                unit_price_col = col_idx
            elif col_type == ColumnType.TOTAL_PRICE:
                total_price_col = col_idx
        
        if not all([quantity_col is not None, unit_price_col is not None, total_price_col is not None]):
            return checks
        
        for row_idx, row in enumerate(sheet_data):
            try:
                # Extract values with proper null checks
                quantity = None
                unit_price = None
                total_price = None
                
                if quantity_col is not None and quantity_col < len(row):
                    quantity = self._parse_number(row[quantity_col])
                if unit_price_col is not None and unit_price_col < len(row):
                    unit_price = self._parse_number(row[unit_price_col])
                if total_price_col is not None and total_price_col < len(row):
                    total_price = self._parse_number(row[total_price_col])
                
                if all(v is not None for v in [quantity, unit_price, total_price]):
                    # Type assertion to help type checker
                    assert quantity is not None and unit_price is not None and total_price is not None
                    calculated_total = quantity * unit_price
                    difference = abs(calculated_total - total_price)
                    tolerance = total_price * self.tolerance_percentage
                    is_valid = difference <= tolerance
                    
                    checks.append(MathematicalCheck(
                        row_index=row_idx,
                        quantity=quantity,
                        unit_price=unit_price,
                        total_price=total_price,
                        calculated_total=calculated_total,
                        difference=difference,
                        tolerance=tolerance,
                        is_valid=is_valid
                    ))
                    
            except (ValueError, IndexError) as e:
                logger.warning(f"Error validating row {row_idx}: {e}")
                continue
        
        return checks
    
    def validate_sheet(self, sheet_data: List[List[str]], 
                      column_mapping: Dict[int, ColumnType],
                      row_classifications: Dict[int, str]) -> ValidationResult:
        """
        Validate a complete sheet with all validation types
        
        Args:
            sheet_data: Sheet data as list of rows
            column_mapping: Dictionary mapping column index to ColumnType
            row_classifications: Dictionary mapping row index to classification
            
        Returns:
            ValidationResult with all validation issues and summary
        """
        issues = []
        
        # Mathematical consistency validation
        math_checks = self.validate_mathematical_consistency(sheet_data, column_mapping)
        for check in math_checks:
            if not check.is_valid:
                issues.append(ValidationIssue(
                    row_index=check.row_index,
                    column_index=None,
                    validation_type=ValidationType.MATHEMATICAL,
                    level=ValidationLevel.ERROR,
                    message=f"Mathematical inconsistency: calculated {check.calculated_total}, actual {check.total_price}",
                    expected_value=check.calculated_total,
                    actual_value=check.total_price,
                    suggestion="Check quantity and unit price calculations"
                ))
        
        # Data type validation
        data_type_issues = self.validate_data_types(sheet_data, column_mapping)
        issues.extend(data_type_issues)
        
        # Business rule validation
        business_issues = self.validate_business_rules(sheet_data, column_mapping, row_classifications)
        issues.extend(business_issues)
        
        # Consistency validation
        consistency_issues = self.validate_consistency(sheet_data, column_mapping)
        issues.extend(consistency_issues)
        
        # Calculate summary
        summary = self._calculate_summary(issues)
        
        # Calculate overall score
        overall_score = self._calculate_overall_score(issues, len(sheet_data))
        
        # Generate confidence factors
        confidence_factors = self._calculate_confidence_factors(issues, sheet_data)
        
        # Generate suggestions
        suggestions = self._generate_suggestions(issues)
        
        return ValidationResult(
            issues=issues,
            summary=summary,
            overall_score=overall_score,
            confidence_factors=confidence_factors,
            suggestions=suggestions
        )
    
    def validate_data_types(self, sheet_data: List[List[str]], 
                           column_mapping: Dict[int, ColumnType]) -> List[ValidationIssue]:
        """Validate data types for each column"""
        issues = []
        
        for col_idx, col_type in column_mapping.items():
            for row_idx, row in enumerate(sheet_data):
                if col_idx >= len(row):
                    continue
                    
                value = row[col_idx]
                if not value or value.strip() == "":
                    continue
                
                if col_type == ColumnType.QUANTITY:
                    if not self._is_valid_number(value):
                        issues.append(ValidationIssue(
                            row_index=row_idx,
                            column_index=col_idx,
                            validation_type=ValidationType.DATA_TYPE,
                            level=ValidationLevel.ERROR,
                            message=f"Invalid quantity format: {value}",
                            expected_value="Numeric value",
                            actual_value=value,
                            suggestion="Enter a valid numeric quantity"
                        ))
                
                elif col_type == ColumnType.UNIT_PRICE:
                    if not self._is_valid_currency(value):
                        issues.append(ValidationIssue(
                            row_index=row_idx,
                            column_index=col_idx,
                            validation_type=ValidationType.DATA_TYPE,
                            level=ValidationLevel.ERROR,
                            message=f"Invalid unit price format: {value}",
                            expected_value="Currency value",
                            actual_value=value,
                            suggestion="Enter a valid currency amount"
                        ))
                
                elif col_type == ColumnType.UNIT:
                    if not self._is_valid_unit(value):
                        issues.append(ValidationIssue(
                            row_index=row_idx,
                            column_index=col_idx,
                            validation_type=ValidationType.DATA_TYPE,
                            level=ValidationLevel.WARNING,
                            message=f"Unusual unit format: {value}",
                            expected_value="Standard unit",
                            actual_value=value,
                            suggestion="Use standard unit formats (m², kg, pcs, etc.)"
                        ))
        
        return issues
    
    def validate_business_rules(self, sheet_data: List[List[str]], 
                              column_mapping: Dict[int, ColumnType],
                              row_classifications: Dict[int, str]) -> List[ValidationIssue]:
        """Validate business rules and logic"""
        issues = []
        
        # Define required column types
        required_types = [ColumnType.DESCRIPTION, ColumnType.QUANTITY, ColumnType.UNIT_PRICE, 
                         ColumnType.TOTAL_PRICE, ColumnType.UNIT, ColumnType.CODE]
        
        # Find required columns
        required_columns = {}
        for col_idx, col_type in column_mapping.items():
            if col_type in required_types:
                required_columns[col_type] = col_idx
        
        # Check for missing required columns in line items
        for row_idx, row in enumerate(sheet_data):
            # Only check line items (not headers, subtotals, etc.)
            row_type = row_classifications.get(row_idx, "")
            if row_type == "primary_line_item":
                missing_columns = []
                for required_type in required_types:
                    if required_type not in required_columns:
                        missing_columns.append(required_type.value)
                    else:
                        col_idx = required_columns[required_type]
                        if col_idx >= len(row) or not row[col_idx] or row[col_idx].strip() == "":
                            missing_columns.append(required_type.value)
                
                if missing_columns:
                    issues.append(ValidationIssue(
                        row_index=row_idx,
                        column_index=None,
                        validation_type=ValidationType.BUSINESS_RULE,
                        level=ValidationLevel.ERROR,
                        message=f"Missing required columns: {', '.join(missing_columns)}",
                        expected_value="All required columns present",
                        actual_value=f"Missing: {', '.join(missing_columns)}",
                        suggestion="Ensure all required columns (Description, Quantity, Unit Price, Total Price, Unit, Code) have values"
                    ))
        
        # Check for negative quantities
        quantity_col = None
        for col_idx, col_type in column_mapping.items():
            if col_type == ColumnType.QUANTITY:
                quantity_col = col_idx
                break
        
        if quantity_col is not None:
            for row_idx, row in enumerate(sheet_data):
                if quantity_col < len(row):
                    try:
                        quantity = self._parse_number(row[quantity_col])
                        if quantity is not None and quantity < 0:
                            issues.append(ValidationIssue(
                                row_index=row_idx,
                                column_index=quantity_col,
                                validation_type=ValidationType.BUSINESS_RULE,
                                level=ValidationLevel.ERROR,
                                message="Negative quantity detected",
                                expected_value="Positive number",
                                actual_value=quantity,
                                suggestion="Quantities should be positive"
                            ))
                    except (ValueError, IndexError):
                        continue
        
        return issues
    
    def validate_consistency(self, sheet_data: List[List[str]], 
                           column_mapping: Dict[int, ColumnType]) -> List[ValidationIssue]:
        """Validate data consistency across the sheet"""
        issues = []
        
        # Check for duplicate descriptions
        description_col = None
        for col_idx, col_type in column_mapping.items():
            if col_type == ColumnType.DESCRIPTION:
                description_col = col_idx
                break
        
        if description_col is not None:
            descriptions = {}
            for row_idx, row in enumerate(sheet_data):
                if description_col < len(row):
                    desc = row[description_col].strip()
                    if desc and desc in descriptions:
                        issues.append(ValidationIssue(
                            row_index=row_idx,
                            column_index=description_col,
                            validation_type=ValidationType.CONSISTENCY,
                            level=ValidationLevel.WARNING,
                            message=f"Duplicate description: {desc}",
                            expected_value="Unique description",
                            actual_value=desc,
                            suggestion="Consider merging duplicate items or adding distinguishing details"
                        ))
                    descriptions[desc] = row_idx
        
        return issues
    
    def _parse_number(self, value: str) -> Optional[float]:
        """Parse a string value to a number"""
        if not value or not value.strip():
            return None
        
        # Remove currency symbols and commas
        cleaned = re.sub(r'[^\d.-]', '', value.strip())
        
        try:
            return float(cleaned)
        except ValueError:
            return None
    
    def _is_valid_number(self, value: str) -> bool:
        """Check if value is a valid number"""
        return self._parse_number(value) is not None
    
    def _is_valid_currency(self, value: str) -> bool:
        """Check if value is a valid currency format"""
        value = value.strip()
        for pattern in self.currency_patterns:
            if re.match(pattern, value):
                return True
        return self._is_valid_number(value)
    
    def _is_valid_unit(self, value: str) -> bool:
        """Check if value is a valid unit format"""
        value = value.strip().lower()
        for pattern in self.unit_patterns:
            if re.match(pattern, value):
                return True
        return True  # Allow custom units with warning
    
    def _calculate_summary(self, issues: List[ValidationIssue]) -> Dict[ValidationLevel, int]:
        """Calculate summary of validation issues by level"""
        summary = {level: 0 for level in ValidationLevel}
        for issue in issues:
            summary[issue.level] += 1
        return summary
    
    def _calculate_overall_score(self, issues: List[ValidationIssue], total_rows: int) -> float:
        """Calculate overall validation score (0-100)"""
        if total_rows == 0:
            return 100.0
        
        error_count = sum(1 for issue in issues if issue.level == ValidationLevel.ERROR)
        warning_count = sum(1 for issue in issues if issue.level == ValidationLevel.WARNING)
        
        # Score calculation: 100 - (errors * 10) - (warnings * 2)
        score = 100.0 - (error_count * 10) - (warning_count * 2)
        return max(0.0, min(100.0, score))
    
    def _calculate_confidence_factors(self, issues: List[ValidationIssue], 
                                    sheet_data: List[List[str]]) -> Dict[str, float]:
        """Calculate confidence factors for different aspects"""
        total_rows = len(sheet_data)
        if total_rows == 0:
            return {"overall": 0.0, "mathematical": 0.0, "data_types": 0.0}
        
        error_count = sum(1 for issue in issues if issue.level == ValidationLevel.ERROR)
        warning_count = sum(1 for issue in issues if issue.level == ValidationLevel.WARNING)
        
        overall_confidence = max(0.0, 1.0 - (error_count / total_rows) - (warning_count / total_rows * 0.5))
        
        return {
            "overall": overall_confidence,
            "mathematical": 1.0 - (sum(1 for issue in issues if issue.validation_type == ValidationType.MATHEMATICAL) / total_rows),
            "data_types": 1.0 - (sum(1 for issue in issues if issue.validation_type == ValidationType.DATA_TYPE) / total_rows)
        }
    
    def _generate_suggestions(self, issues: List[ValidationIssue]) -> List[str]:
        """Generate improvement suggestions based on validation issues"""
        suggestions = []
        
        error_count = sum(1 for issue in issues if issue.level == ValidationLevel.ERROR)
        warning_count = sum(1 for issue in issues if issue.level == ValidationLevel.WARNING)
        
        if error_count > 0:
            suggestions.append(f"Fix {error_count} critical validation errors")
        
        if warning_count > 0:
            suggestions.append(f"Review {warning_count} validation warnings")
        
        math_issues = [issue for issue in issues if issue.validation_type == ValidationType.MATHEMATICAL]
        if math_issues:
            suggestions.append("Review mathematical calculations for consistency")
        
        data_type_issues = [issue for issue in issues if issue.validation_type == ValidationType.DATA_TYPE]
        if data_type_issues:
            suggestions.append("Standardize data formats across columns")
        
        return suggestions
 