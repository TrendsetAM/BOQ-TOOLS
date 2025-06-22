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
 