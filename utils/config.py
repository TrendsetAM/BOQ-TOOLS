"""
BOQ Tools Configuration System
Comprehensive configuration for Bill of Quantities Excel processing
"""

from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum
import logging
import os
import json

try:
    from appdirs import user_config_dir
except ImportError:
    user_config_dir = None

logger = logging.getLogger(__name__)


class ColumnType(Enum):
    """Enumeration for different column types in BoQ sheets"""
    DESCRIPTION = "description"
    QUANTITY = "quantity"
    UNIT_PRICE = "unit_price"
    TOTAL_PRICE = "total_price"
    UNIT = "unit"
    CODE = "code"
    IGNORE = "ignore"


@dataclass
class ColumnMapping:
    """Configuration for column mapping with confidence scoring"""
    keywords: List[str]
    weight: float = 1.0
    required: bool = False
    data_type: str = "text"
    validation_pattern: Optional[str] = None


@dataclass
class SheetClassification:
    """Configuration for sheet classification and scoring"""
    keywords: List[str]
    weight: float = 1.0
    min_confidence: float = 0.6
    sheet_type: str = "unknown"


@dataclass
class ValidationThresholds:
    """Validation thresholds and confidence scores"""
    min_column_confidence: float = 0.7
    min_sheet_confidence: float = 0.6
    max_empty_rows_percentage: float = 0.3
    min_data_rows: int = 5
    max_header_rows: int = 10


@dataclass
class ProcessingLimits:
    """File processing limits and performance settings"""
    max_file_size_mb: int = 50
    max_sheets_per_file: int = 20
    max_rows_per_sheet: int = 10000
    max_columns_per_sheet: int = 50
    timeout_seconds: int = 300
    memory_limit_mb: int = 512


@dataclass
class ExportSettings:
    """Default export settings and formats"""
    output_format: str = "xlsx"
    include_summary: bool = True
    include_validation_report: bool = True
    sheet_name_template: str = "Processed_{original_name}"
    backup_original: bool = True
    compression_level: int = 6


class BOQConfig:
    """
    Comprehensive configuration system for BOQ processing
    """
    
    def __init__(self):
        self._setup_column_mappings()
        self._setup_sheet_classifications()
        self._setup_row_patterns()
        self._setup_validation_thresholds()
        self._setup_processing_limits()
        self._setup_export_settings()
    
    def _setup_column_mappings(self) -> None:
        """Initialize column mapping configurations"""
        self.column_mappings: Dict[ColumnType, ColumnMapping] = {
            ColumnType.DESCRIPTION: ColumnMapping(
                keywords=["description", "item", "work", "activity", "task", "detail", 
                         "particulars", "scope", "specification", "work description"],
                weight=1.2,
                required=True,
                data_type="text"
            ),
            
            ColumnType.QUANTITY: ColumnMapping(
                keywords=["qty", "quantity", "no", "number", "count",
                         "qty.", "qty:", "quantity:", "nos", "numbers"],
                weight=1.1,
                required=True,
                data_type="numeric",
                validation_pattern=r"^\d+(\.\d+)?$"
            ),
            
            ColumnType.UNIT_PRICE: ColumnMapping(
                keywords=["unit price", "rate", "price per unit", "unit cost", "unit rate",
                         "rate per unit", "price/unit", "unit price:", "rate:", "cost per unit"],
                weight=1.0,
                required=True,
                data_type="currency",
                validation_pattern=r"^\d+(\.\d{1,2})?$"
            ),
            
            ColumnType.TOTAL_PRICE: ColumnMapping(
                keywords=["total", "total price", "total cost", "value",
                         "total amount", "sum", "total value", "cost", "price"],
                weight=1.0,
                required=True,
                data_type="currency",
                validation_pattern=r"^\d+(\.\d{1,2})?$"
            ),
            
            ColumnType.UNIT: ColumnMapping(
                keywords=["unit", "measurement", "measure", "uom", "unit of measure",
                         "unit:", "measurement unit"],
                weight=0.9,
                required=True,
                data_type="text"
            ),
            
            ColumnType.CODE: ColumnMapping(
                keywords=["code", "item code", "reference", "ref", "item no", "item number",
                         "code:", "reference:", "item code:", "boq code"],
                weight=0.7,
                required=True,
                data_type="text"
            )
        }
    
    def _setup_sheet_classifications(self) -> None:
        """Initialize sheet classification configurations"""
        self.sheet_classifications: List[SheetClassification] = [
            SheetClassification(
                keywords=["bill of quantities", "boq", "bill of quantity", "quantity survey",
                         "take off", "takeoff", "measurement", "quantities"],
                weight=1.5,
                min_confidence=0.7,
                sheet_type="boq_main"
            ),
            
            SheetClassification(
                keywords=["summary", "total", "summary sheet", "total sheet", "summary of",
                         "total summary", "grand total", "final total"],
                weight=1.2,
                min_confidence=0.6,
                sheet_type="summary"
            ),
            
            SheetClassification(
                keywords=["preliminaries", "preliminary", "prelim", "general items",
                         "general requirements", "preliminary items"],
                weight=1.1,
                min_confidence=0.6,
                sheet_type="preliminaries"
            ),
            
            SheetClassification(
                keywords=["substructure", "foundation", "excavation", "concrete",
                         "reinforcement", "formwork", "substructure works"],
                weight=1.0,
                min_confidence=0.6,
                sheet_type="substructure"
            ),
            
            SheetClassification(
                keywords=["superstructure", "structure", "framing", "roof", "walls",
                         "superstructure works", "structural works"],
                weight=1.0,
                min_confidence=0.6,
                sheet_type="superstructure"
            ),
            
            SheetClassification(
                keywords=["finishes", "finishing", "paint", "tiles", "flooring",
                         "finishing works", "interior finishes"],
                weight=1.0,
                min_confidence=0.6,
                sheet_type="finishes"
            ),
            
            SheetClassification(
                keywords=["services", "mep", "mechanical", "electrical", "plumbing",
                         "services works", "m&e", "building services"],
                weight=1.0,
                min_confidence=0.6,
                sheet_type="services"
            ),
            
            SheetClassification(
                keywords=["external works", "landscaping", "drainage", "roads",
                         "external works", "site works"],
                weight=1.0,
                min_confidence=0.6,
                sheet_type="external_works"
            )
        ]
    
    def _setup_row_patterns(self) -> None:
        """Initialize row classification patterns"""
        self.row_patterns = {
            "subtotal_indicators": [
                "subtotal", "sub total", "sub-total", "sub total:", "subtotal:",
                "total for", "total of", "section total", "group total"
            ],
            
            "section_headers": [
                "section", "division", "group", "category", "part",
                "section:", "division:", "group:", "category:"
            ],
            
            "page_breaks": [
                "page", "sheet", "continued", "cont'd", "continued on",
                "page break", "new page"
            ],
            
            "summary_rows": [
                "grand total", "final total", "total project", "project total",
                "overall total", "total amount", "final amount"
            ],
            
            "exclude_patterns": [
                "blank", "empty", "n/a", "not applicable", "tbd", "to be determined",
                "pending", "under review", "draft"
            ]
        }
    
    def _setup_validation_thresholds(self) -> None:
        """Initialize validation thresholds"""
        self.validation_thresholds = ValidationThresholds(
            min_column_confidence=0.7,
            min_sheet_confidence=0.6,
            max_empty_rows_percentage=0.3,
            min_data_rows=5,
            max_header_rows=10
        )
    
    def _setup_processing_limits(self) -> None:
        """Initialize processing limits"""
        self.processing_limits = ProcessingLimits(
            max_file_size_mb=50,
            max_sheets_per_file=20,
            max_rows_per_sheet=10000,
            max_columns_per_sheet=50,
            timeout_seconds=300,
            memory_limit_mb=512
        )
    
    def _setup_export_settings(self) -> None:
        """Initialize export settings"""
        self.export_settings = ExportSettings(
            output_format="xlsx",
            include_summary=True,
            include_validation_report=True,
            sheet_name_template="Processed_{original_name}",
            backup_original=True,
            compression_level=6
        )
    
    def get_column_mapping(self, column_type: ColumnType) -> Optional[ColumnMapping]:
        """Get column mapping configuration for a specific type"""
        return self.column_mappings.get(column_type)
    
    def get_all_column_types(self) -> List[ColumnType]:
        """Get all available column types"""
        return list(self.column_mappings.keys())
    
    def get_required_columns(self) -> List[ColumnType]:
        """Get list of required column types"""
        return [col_type for col_type, mapping in self.column_mappings.items() 
                if mapping.required]
    
    def get_sheet_classification(self, sheet_name: str, content: List[str]) -> Tuple[str, float]:
        """
        Classify a sheet based on its name and content
        Returns: (sheet_type, confidence_score)
        """
        best_match = ("unknown", 0.0)
        
        for classification in self.sheet_classifications:
            score = self._calculate_sheet_score(sheet_name, content, classification)
            if score > best_match[1] and score >= classification.min_confidence:
                best_match = (classification.sheet_type, score)
        
        return best_match
    
    def _calculate_sheet_score(self, sheet_name: str, content: List[str], 
                              classification: SheetClassification) -> float:
        """Calculate confidence score for sheet classification"""
        score = 0.0
        total_weight = 0.0
        
        # Check sheet name
        sheet_name_lower = sheet_name.lower()
        for keyword in classification.keywords:
            if keyword.lower() in sheet_name_lower:
                score += classification.weight * 2  # Higher weight for name matches
                total_weight += classification.weight * 2
        
        # Check content
        content_text = " ".join(content).lower()
        for keyword in classification.keywords:
            if keyword.lower() in content_text:
                score += classification.weight
                total_weight += classification.weight
        
        return score / total_weight if total_weight > 0 else 0.0
    
    def validate_configuration(self) -> List[str]:
        """Validate the configuration and return any issues"""
        issues = []
        
        # Check for duplicate keywords in column mappings
        all_keywords = []
        for col_type, mapping in self.column_mappings.items():
            for keyword in mapping.keywords:
                if keyword in all_keywords:
                    issues.append(f"Duplicate keyword '{keyword}' in column mappings")
                all_keywords.append(keyword)
        
        # Check validation thresholds
        if self.validation_thresholds.min_column_confidence > 1.0:
            issues.append("min_column_confidence cannot be greater than 1.0")
        
        if self.validation_thresholds.max_empty_rows_percentage > 1.0:
            issues.append("max_empty_rows_percentage cannot be greater than 1.0")
        
        # Check processing limits
        if self.processing_limits.max_file_size_mb <= 0:
            issues.append("max_file_size_mb must be positive")
        
        if self.processing_limits.timeout_seconds <= 0:
            issues.append("timeout_seconds must be positive")
        
        return issues
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary for serialization"""
        return {
            "column_mappings": {
                col_type.value: {
                    "keywords": mapping.keywords,
                    "weight": mapping.weight,
                    "required": mapping.required,
                    "data_type": mapping.data_type,
                    "validation_pattern": mapping.validation_pattern
                }
                for col_type, mapping in self.column_mappings.items()
            },
            "validation_thresholds": {
                "min_column_confidence": self.validation_thresholds.min_column_confidence,
                "min_sheet_confidence": self.validation_thresholds.min_sheet_confidence,
                "max_empty_rows_percentage": self.validation_thresholds.max_empty_rows_percentage,
                "min_data_rows": self.validation_thresholds.min_data_rows,
                "max_header_rows": self.validation_thresholds.max_header_rows
            },
            "processing_limits": {
                "max_file_size_mb": self.processing_limits.max_file_size_mb,
                "max_sheets_per_file": self.processing_limits.max_sheets_per_file,
                "max_rows_per_sheet": self.processing_limits.max_rows_per_sheet,
                "max_columns_per_sheet": self.processing_limits.max_columns_per_sheet,
                "timeout_seconds": self.processing_limits.timeout_seconds,
                "memory_limit_mb": self.processing_limits.memory_limit_mb
            },
            "export_settings": {
                "output_format": self.export_settings.output_format,
                "include_summary": self.export_settings.include_summary,
                "include_validation_report": self.export_settings.include_validation_report,
                "sheet_name_template": self.export_settings.sheet_name_template,
                "backup_original": self.export_settings.backup_original,
                "compression_level": self.export_settings.compression_level
            }
        }


# Global configuration instance
config = BOQConfig()


def get_config() -> BOQConfig:
    """Get the global configuration instance"""
    return config


def validate_and_log_config() -> bool:
    """Validate configuration and log any issues"""
    issues = config.validate_configuration()
    
    if issues:
        logger.error("Configuration validation failed:")
        for issue in issues:
            logger.error(f"  - {issue}")
        return False
    else:
        logger.info("Configuration validation passed")
        return True


def get_user_config_path(filename: str) -> str:
    app_name = "BOQ-TOOLS"
    if user_config_dir:
        config_dir = user_config_dir(app_name)
    else:
        if os.name == 'nt':
            config_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), app_name)
        else:
            config_dir = os.path.join(os.path.expanduser('~/.config'), app_name)
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, filename)


def ensure_default_config(filename: str, default_path: str, default_data: dict = None):
    user_path = get_user_config_path(filename)
    if not os.path.exists(user_path):
        if os.path.exists(default_path):
            with open(default_path, 'r', encoding='utf-8') as fsrc:
                data = fsrc.read()
            with open(user_path, 'w', encoding='utf-8') as fdst:
                fdst.write(data)
        elif default_data is not None:
            with open(user_path, 'w', encoding='utf-8') as fdst:
                json.dump(default_data, fdst, indent=2, ensure_ascii=False)
    return user_path


if __name__ == "__main__":
    # Test configuration
    logging.basicConfig(level=logging.INFO)
    validate_and_log_config()
    
    # Print configuration summary
    print(f"Column types: {len(config.get_all_column_types())}")
    print(f"Required columns: {len(config.get_required_columns())}")
    print(f"Sheet classifications: {len(config.sheet_classifications)}")
    print(f"Row patterns: {len(config.row_patterns)}") 