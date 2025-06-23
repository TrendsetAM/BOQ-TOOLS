"""
Mapping Generator for BOQ Tools
Unified mapping structure generation with comprehensive confidence scoring
"""

import logging
import json
from typing import Dict, List, Tuple, Optional, Any, Set
from dataclasses import dataclass, asdict
from datetime import datetime
from enum import Enum
from pathlib import Path

from utils.config import get_config, ColumnType

logger = logging.getLogger(__name__)


class ProcessingStatus(Enum):
    """Processing status enumeration"""
    SUCCESS = "success"
    PARTIAL = "partial"
    FAILED = "failed"
    NEEDS_REVIEW = "needs_review"


class ReviewFlag(Enum):
    """Review flag types"""
    LOW_CONFIDENCE = "low_confidence"
    VALIDATION_ERRORS = "validation_errors"
    AMBIGUOUS_MAPPING = "ambiguous_mapping"
    MISSING_DATA = "missing_data"
    INCONSISTENT_FORMAT = "inconsistent_format"
    MANUAL_REVIEW_REQUIRED = "manual_review_required"


@dataclass
class ColumnMappingInfo:
    """Detailed column mapping information"""
    column_index: int
    original_header: str
    normalized_header: str
    mapped_type: str
    confidence: float
    alternatives: List[Dict[str, Any]]
    reasoning: List[str]
    is_required: bool
    validation_status: str
    is_user_edited: bool


@dataclass
class RowClassificationInfo:
    """Detailed row classification information"""
    row_index: int
    row_type: str
    confidence: float
    completeness_score: float
    hierarchical_level: Optional[int]
    section_title: Optional[str]
    validation_errors: List[str]
    reasoning: List[str]


@dataclass
class ValidationSummary:
    """Validation summary for a sheet"""
    overall_score: float
    mathematical_consistency: float
    data_type_quality: float
    business_rule_compliance: float
    error_count: int
    warning_count: int
    info_count: int
    suggestions: List[str]


@dataclass
class SheetMapping:
    """Complete mapping information for a single sheet"""
    sheet_name: str
    processing_status: ProcessingStatus
    row_count: int
    column_count: int
    header_row_index: int
    header_confidence: float
    column_mappings: List[ColumnMappingInfo]
    row_classifications: List[RowClassificationInfo]
    validation_summary: ValidationSummary
    overall_confidence: float
    column_mapping_confidence: float
    row_classification_confidence: float
    data_quality_confidence: float
    review_flags: List[ReviewFlag]
    manual_review_items: List[Dict[str, Any]]
    processing_notes: List[str]
    warnings: List[str]
    processing_time: float
    sheet_type: str  # User or classifier assigned type (e.g., 'BOQ', 'Info', 'Ignore')


@dataclass
class FileMetadata:
    """File metadata information"""
    filename: str
    file_path: str
    file_size_mb: float
    file_format: str
    processing_date: datetime
    total_sheets: int
    visible_sheets: int
    processing_version: str


@dataclass
class ProcessingSummary:
    """Overall processing summary"""
    total_rows_processed: int
    total_columns_mapped: int
    successful_sheets: int
    partial_sheets: int
    failed_sheets: int
    sheets_needing_review: int
    
    # Quality metrics
    average_confidence: float
    average_data_quality: float
    total_validation_errors: int
    total_validation_warnings: int
    
    # Processing statistics
    processing_time_total: float
    processing_notes: List[str]
    recommendations: List[str]


@dataclass
class FileMapping:
    """Complete file mapping structure"""
    metadata: FileMetadata
    sheets: List[SheetMapping]
    global_confidence: float
    processing_summary: ProcessingSummary
    review_flags: List[ReviewFlag]
    export_ready: bool


class MappingGenerator:
    """
    Unified mapping structure generator with comprehensive confidence scoring
    """
    
    def __init__(self, processing_version: str = "1.0.0"):
        """
        Initialize the mapping generator
        
        Args:
            processing_version: Version of the processing pipeline
        """
        self.config = get_config()
        self.processing_version = processing_version
        self.review_thresholds = {
            'low_confidence': 0.6,
            'validation_errors': 5,
            'ambiguous_mapping': 0.7,
            'missing_data': 0.3
        }
        
        logger.info("Mapping Generator initialized")
    
    def generate_file_mapping(self, processor_results: Dict[str, Any]) -> FileMapping:
        """
        Generate complete file mapping structure
        
        Args:
            processor_results: Dictionary containing all processing results
            
        Returns:
            Complete FileMapping structure
        """
        logger.info("Generating unified file mapping structure")
        
        try:
            # Extract basic information
            file_info = processor_results.get('file_info', {})
            sheet_data = processor_results.get('sheet_data', {})
            column_mappings = processor_results.get('column_mappings', {})
            row_classifications = processor_results.get('row_classifications', {})
            validation_results = processor_results.get('validation_results', {})
            sheet_types = processor_results.get('sheet_types', {})
            
            # Create metadata
            metadata = self._create_file_metadata(file_info)
            
            # Process each sheet
            sheet_mappings = []
            for sheet_name in sheet_data.keys():
                sheet_type = sheet_types.get(sheet_name, 'BOQ')
                sheet_mapping = self._create_sheet_mapping(
                    sheet_name=sheet_name,
                    sheet_data=sheet_data.get(sheet_name, []),
                    column_mapping=column_mappings.get(sheet_name, {}),
                    row_classification=row_classifications.get(sheet_name, {}),
                    validation_result=validation_results.get(sheet_name, {}),
                    sheet_type=sheet_type
                )
                sheet_mappings.append(sheet_mapping)
            
            # Calculate global confidence
            global_confidence = self.calculate_global_confidence(sheet_mappings)
            
            # Create processing summary
            processing_summary = self.create_processing_summary(sheet_mappings)
            
            # Flag items needing manual review
            review_flags = self._identify_global_review_flags(sheet_mappings)
            
            # Determine if export is ready
            export_ready = self._is_export_ready(sheet_mappings, global_confidence)
            
            file_mapping = FileMapping(
                metadata=metadata,
                sheets=sheet_mappings,
                global_confidence=global_confidence,
                processing_summary=processing_summary,
                review_flags=review_flags,
                export_ready=export_ready
            )
            
            logger.info(f"File mapping generated successfully: {len(sheet_mappings)} sheets, "
                       f"global confidence: {global_confidence:.2f}")
            
            return file_mapping
            
        except Exception as e:
            logger.error(f"Error generating file mapping: {e}")
            raise
    
    def calculate_global_confidence(self, sheet_mappings: List[SheetMapping]) -> float:
        """
        Calculate global confidence score across all sheets
        
        Args:
            sheet_mappings: List of sheet mappings
            
        Returns:
            Global confidence score
        """
        if not sheet_mappings:
            return 0.0
        
        # Weight factors for different confidence types
        weights = {
            'overall_confidence': 0.4,
            'column_mapping_confidence': 0.3,
            'row_classification_confidence': 0.2,
            'data_quality_confidence': 0.1
        }
        
        total_confidence = 0.0
        total_weight = 0.0
        
        for sheet in sheet_mappings:
            # Calculate weighted confidence for this sheet
            sheet_confidence = (
                sheet.overall_confidence * weights['overall_confidence'] +
                sheet.column_mapping_confidence * weights['column_mapping_confidence'] +
                sheet.row_classification_confidence * weights['row_classification_confidence'] +
                sheet.data_quality_confidence * weights['data_quality_confidence']
            )
            
            # Weight by sheet importance (more rows = more important)
            sheet_weight = min(sheet.row_count / 100.0, 1.0)  # Cap at 1.0
            
            total_confidence += sheet_confidence * sheet_weight
            total_weight += sheet_weight
        
        return total_confidence / total_weight if total_weight > 0 else 0.0
    
    def create_processing_summary(self, sheet_mappings: List[SheetMapping]) -> ProcessingSummary:
        """
        Create comprehensive processing summary
        
        Args:
            sheet_mappings: List of sheet mappings
            
        Returns:
            ProcessingSummary with detailed statistics
        """
        total_rows = sum(sheet.row_count for sheet in sheet_mappings)
        total_columns = sum(len(sheet.column_mappings) for sheet in sheet_mappings)
        
        # Count sheets by status
        successful_sheets = sum(1 for s in sheet_mappings if s.processing_status == ProcessingStatus.SUCCESS)
        partial_sheets = sum(1 for s in sheet_mappings if s.processing_status == ProcessingStatus.PARTIAL)
        failed_sheets = sum(1 for s in sheet_mappings if s.processing_status == ProcessingStatus.FAILED)
        review_sheets = sum(1 for s in sheet_mappings if s.processing_status == ProcessingStatus.NEEDS_REVIEW)
        
        # Calculate averages
        avg_confidence = sum(s.overall_confidence for s in sheet_mappings) / len(sheet_mappings) if sheet_mappings else 0.0
        avg_data_quality = sum(s.data_quality_confidence for s in sheet_mappings) / len(sheet_mappings) if sheet_mappings else 0.0
        
        # Count validation issues
        total_errors = sum(s.validation_summary.error_count for s in sheet_mappings)
        total_warnings = sum(s.validation_summary.warning_count for s in sheet_mappings)
        
        # Calculate total processing time
        total_processing_time = sum(s.processing_time for s in sheet_mappings)
        
        # Generate recommendations
        recommendations = self._generate_recommendations(sheet_mappings)
        
        # Collect processing notes
        all_notes = []
        for sheet in sheet_mappings:
            all_notes.extend(sheet.processing_notes)
        
        return ProcessingSummary(
            total_rows_processed=total_rows,
            total_columns_mapped=total_columns,
            successful_sheets=successful_sheets,
            partial_sheets=partial_sheets,
            failed_sheets=failed_sheets,
            sheets_needing_review=review_sheets,
            average_confidence=avg_confidence,
            average_data_quality=avg_data_quality,
            total_validation_errors=total_errors,
            total_validation_warnings=total_warnings,
            processing_time_total=total_processing_time,
            processing_notes=all_notes,
            recommendations=recommendations
        )
    
    def flag_manual_review_items(self, sheet_mappings: List[SheetMapping]) -> List[Dict[str, Any]]:
        """
        Flag items that need manual review
        
        Args:
            sheet_mappings: List of sheet mappings
            
        Returns:
            List of items requiring manual review
        """
        review_items = []
        
        for sheet in sheet_mappings:
            # Check for low confidence mappings
            for col_mapping in sheet.column_mappings:
                if col_mapping.confidence < self.review_thresholds['low_confidence']:
                    review_items.append({
                        'type': 'low_confidence_mapping',
                        'sheet_name': sheet.sheet_name,
                        'row_index': None,
                        'column_index': col_mapping.column_index,
                        'header': col_mapping.original_header,
                        'mapped_type': col_mapping.mapped_type,
                        'confidence': col_mapping.confidence,
                        'suggestion': 'Review column mapping manually'
                    })
            
            # Check for validation errors
            if sheet.validation_summary.error_count > self.review_thresholds['validation_errors']:
                review_items.append({
                    'type': 'validation_errors',
                    'sheet_name': sheet.sheet_name,
                    'row_index': None,
                    'column_index': None,
                    'error_count': sheet.validation_summary.error_count,
                    'suggestion': 'Review and fix validation errors'
                })
            
            # Check for ambiguous mappings
            ambiguous_mappings = [cm for cm in sheet.column_mappings 
                                if len(cm.alternatives) > 1 and cm.confidence < self.review_thresholds['ambiguous_mapping']]
            if ambiguous_mappings:
                review_items.append({
                    'type': 'ambiguous_mappings',
                    'sheet_name': sheet.sheet_name,
                    'mappings': [
                        {
                            'column_index': cm.column_index,
                            'header': cm.original_header,
                            'alternatives': cm.alternatives
                        }
                        for cm in ambiguous_mappings
                    ],
                    'suggestion': 'Review ambiguous column mappings'
                })
            
            # Check for missing data
            missing_data_rows = [rc for rc in sheet.row_classifications 
                               if rc.completeness_score < self.review_thresholds['missing_data']]
            if missing_data_rows:
                review_items.append({
                    'type': 'missing_data',
                    'sheet_name': sheet.sheet_name,
                    'row_count': len(missing_data_rows),
                    'suggestion': 'Review rows with missing required data'
                })
        
        return review_items
    
    def export_mapping_to_json(self, mapping: FileMapping, output_path: Optional[Path] = None) -> str:
        """
        Export mapping structure to JSON format
        
        Args:
            mapping: FileMapping structure to export
            output_path: Optional path to save the JSON file
            
        Returns:
            JSON string representation
        """
        try:
            # Convert to dictionary with proper serialization
            mapping_dict = self._serialize_mapping(mapping)
            
            # Convert to JSON
            json_string = json.dumps(mapping_dict, indent=2, default=str)
            
            # Save to file if path provided
            if output_path:
                output_path = Path(output_path)
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_text(json_string, encoding='utf-8')
                logger.info(f"Mapping exported to: {output_path}")
            
            return json_string
            
        except Exception as e:
            logger.error(f"Error exporting mapping to JSON: {e}")
            raise
    
    def _create_file_metadata(self, file_info: Dict[str, Any]) -> FileMetadata:
        """Create file metadata from file info"""
        return FileMetadata(
            filename=file_info.get('filename', 'unknown'),
            file_path=file_info.get('file_path', ''),
            file_size_mb=file_info.get('file_size_mb', 0.0),
            file_format=file_info.get('file_format', 'unknown'),
            processing_date=datetime.now(),
            total_sheets=file_info.get('total_sheets', 0),
            visible_sheets=file_info.get('visible_sheets', 0),
            processing_version=self.processing_version
        )
    
    def _create_sheet_mapping(self, sheet_name: str, sheet_data: List[List[str]],
                             column_mapping: Any, row_classification: Any,
                             validation_result: Any, sheet_type: str = "BOQ") -> SheetMapping:
        """Create detailed sheet mapping"""
        
        # Handle MappingResult object to get all columns
        all_column_data = {}
        if hasattr(column_mapping, 'header_row'):
            original_headers = column_mapping.header_row.headers
            for i, header in enumerate(original_headers):
                all_column_data[i] = {
                    'column_index': i,
                    'original_header': header,
                    'normalized_header': header.lower().strip(),
                    'mapped_type': 'ignore',
                    'confidence': 0.0,
                    'alternatives': [],
                    'reasoning': ['Not automatically mapped'],
                    'is_user_edited': False,
                }
        
        if hasattr(column_mapping, 'mappings'):
            for col_info in column_mapping.mappings:
                if col_info.column_index in all_column_data:
                    all_column_data[col_info.column_index].update({
                        'normalized_header': col_info.normalized_header,
                        'mapped_type': col_info.mapped_type,
                        'confidence': col_info.confidence,
                        'alternatives': col_info.alternatives,
                        'reasoning': col_info.reasoning,
                    })

        # The column_mapper should have already enforced uniqueness.
        # We just need to set the 'is_required' flag based on the final type.
        required_types_from_config = self.config.get_required_columns()
        required_type_values = {rt.value for rt in required_types_from_config}
        
        for idx, data in all_column_data.items():
            mapped_type_str = data['mapped_type'].value if hasattr(data['mapped_type'], 'value') else str(data['mapped_type'])
            data['is_required'] = mapped_type_str in required_type_values

        # Convert the dictionary of data into a list of ColumnMappingInfo objects
        column_mappings = []
        for idx in sorted(all_column_data.keys()):
            data = all_column_data[idx]
            mapped_type_str = data['mapped_type'].value if hasattr(data['mapped_type'], 'value') else str(data['mapped_type'])
            
            # Alternatives are now expected to be List[Tuple[ColumnType, float]]
            alts_list = []
            if data['alternatives']:
                 alts_list = [{'type': alt[0].value, 'confidence': alt[1]} for alt in data['alternatives']]

            column_mappings.append(ColumnMappingInfo(
                column_index=data['column_index'],
                original_header=data['original_header'],
                normalized_header=data['normalized_header'],
                mapped_type=mapped_type_str,
                confidence=data['confidence'],
                alternatives=alts_list,
                reasoning=data['reasoning'],
                is_required=data['is_required'],
                is_user_edited=data['is_user_edited'],
                validation_status='valid'
            ))
        
        # Extract row classification information
        row_classifications = []
        if hasattr(row_classification, 'classifications'):
            # Handle ClassificationResult object
            for row_info in row_classification.classifications:
                row_class = RowClassificationInfo(
                    row_index=row_info.row_index,
                    row_type=row_info.row_type.value,
                    confidence=row_info.confidence,
                    completeness_score=row_info.completeness_score,
                    hierarchical_level=row_info.hierarchical_level,
                    section_title=row_info.section_title,
                    validation_errors=row_info.validation_errors,
                    reasoning=row_info.reasoning
                )
                row_classifications.append(row_class)
        else:
            # Handle dictionary (fallback)
            for row_info in row_classification.get('classifications', []):
                row_class = RowClassificationInfo(
                    row_index=row_info.get('row_index', 0),
                    row_type=row_info.get('row_type', ''),
                    confidence=row_info.get('confidence', 0.0),
                    completeness_score=row_info.get('completeness_score', 0.0),
                    hierarchical_level=row_info.get('hierarchical_level'),
                    section_title=row_info.get('section_title'),
                    validation_errors=row_info.get('validation_errors', []),
                    reasoning=row_info.get('reasoning', [])
                )
                row_classifications.append(row_class)
        
        # Create validation summary
        if hasattr(validation_result, 'overall_score'):
            # Handle ValidationResult object and normalize score
            raw_score = validation_result.overall_score
            # The validator incorrectly returns a 0-100 score. Normalize to 0-1.
            normalized_score = raw_score / 100.0 if raw_score > 1.0 else raw_score
            
            validation_summary = ValidationSummary(
                overall_score=normalized_score,
                mathematical_consistency=validation_result.confidence_factors.get('mathematical_consistency', 0.0),
                data_type_quality=validation_result.confidence_factors.get('data_type_quality', 0.0),
                business_rule_compliance=validation_result.confidence_factors.get('business_rule_compliance', 0.0),
                error_count=len([i for i in validation_result.issues if i.level == 'error']),
                warning_count=len([i for i in validation_result.issues if i.level == 'warning']),
                info_count=len([i for i in validation_result.issues if i.level == 'info']),
                suggestions=validation_result.suggestions
            )
        else:
            # Handle dictionary (fallback) and normalize
            raw_score = validation_result.get('overall_score', 0.0)
            normalized_score = raw_score / 100.0 if raw_score > 1.0 else raw_score

            validation_summary = ValidationSummary(
                overall_score=normalized_score,
                mathematical_consistency=validation_result.get('mathematical_consistency', 0.0),
                data_type_quality=validation_result.get('data_type_quality', 0.0),
                business_rule_compliance=validation_result.get('business_rule_compliance', 0.0),
                error_count=validation_result.get('error_count', 0),
                warning_count=validation_result.get('warning_count', 0),
                info_count=validation_result.get('info_count', 0),
                suggestions=validation_result.get('suggestions', [])
            )
        
        # Calculate confidence scores
        if hasattr(column_mapping, 'overall_confidence'):
            column_mapping_confidence = column_mapping.overall_confidence
        else:
            column_mapping_confidence = column_mapping.get('overall_confidence', 0.0)
            
        if hasattr(row_classification, 'overall_quality_score'):
            row_classification_confidence = row_classification.overall_quality_score
        else:
            row_classification_confidence = row_classification.get('overall_quality_score', 0.0)
            
        data_quality_confidence = validation_summary.overall_score
        
        # Calculate overall confidence
        overall_confidence = (
            column_mapping_confidence * 0.4 +
            row_classification_confidence * 0.3 +
            data_quality_confidence * 0.3
        )
        
        # Determine processing status
        processing_status = self._determine_processing_status(
            overall_confidence, validation_summary, column_mappings, row_classifications
        )
        
        # Generate review flags
        review_flags = self._generate_sheet_review_flags(
            overall_confidence, validation_summary, column_mappings, row_classifications
        )
        
        # Flag manual review items
        manual_review_items = self._flag_sheet_manual_review_items(
            column_mappings, row_classifications, validation_summary
        )
        
        # Generate processing notes and warnings
        processing_notes, warnings = self._generate_processing_notes(
            column_mapping, row_classification, validation_result
        )
        
        # Get header information
        if hasattr(column_mapping, 'header_row'):
            header_row_index = column_mapping.header_row.row_index
            header_confidence = column_mapping.header_row.confidence
        else:
            header_row_index = column_mapping.get('header_row_index', 0)
            header_confidence = column_mapping.get('header_confidence', 0.0)
        
        return SheetMapping(
            sheet_name=sheet_name,
            processing_status=processing_status,
            row_count=len(sheet_data),
            column_count=len(sheet_data[0]) if sheet_data else 0,
            header_row_index=header_row_index,
            header_confidence=header_confidence,
            column_mappings=column_mappings,
            row_classifications=row_classifications,
            validation_summary=validation_summary,
            overall_confidence=overall_confidence,
            column_mapping_confidence=column_mapping_confidence,
            row_classification_confidence=row_classification_confidence,
            data_quality_confidence=data_quality_confidence,
            review_flags=review_flags,
            manual_review_items=manual_review_items,
            processing_notes=processing_notes,
            warnings=warnings,
            processing_time=0.0,  # Would be calculated during actual processing
            sheet_type=sheet_type
        )
    
    def _determine_processing_status(self, overall_confidence: float, validation_summary: ValidationSummary,
                                   column_mappings: List[ColumnMappingInfo], 
                                   row_classifications: List[RowClassificationInfo]) -> ProcessingStatus:
        """Determine processing status based on results"""
        
        if overall_confidence >= 0.8 and validation_summary.error_count == 0:
            return ProcessingStatus.SUCCESS
        elif overall_confidence >= 0.6 and validation_summary.error_count < 5:
            return ProcessingStatus.PARTIAL
        elif validation_summary.error_count > 10 or overall_confidence < 0.4:
            return ProcessingStatus.FAILED
        else:
            return ProcessingStatus.NEEDS_REVIEW
    
    def _generate_sheet_review_flags(self, overall_confidence: float, validation_summary: ValidationSummary,
                                   column_mappings: List[ColumnMappingInfo], 
                                   row_classifications: List[RowClassificationInfo]) -> List[ReviewFlag]:
        """Generate review flags for a sheet"""
        flags = []
        
        if overall_confidence < self.review_thresholds['low_confidence']:
            flags.append(ReviewFlag.LOW_CONFIDENCE)
        
        if validation_summary.error_count > self.review_thresholds['validation_errors']:
            flags.append(ReviewFlag.VALIDATION_ERRORS)
        
        # Check for ambiguous mappings
        ambiguous_count = sum(1 for cm in column_mappings 
                            if len(cm.alternatives) > 1 and cm.confidence < self.review_thresholds['ambiguous_mapping'])
        if ambiguous_count > 0:
            flags.append(ReviewFlag.AMBIGUOUS_MAPPING)
        
        # Check for missing data
        missing_data_count = sum(1 for rc in row_classifications 
                               if rc.completeness_score < self.review_thresholds['missing_data'])
        if missing_data_count > 0:
            flags.append(ReviewFlag.MISSING_DATA)
        
        if flags:
            flags.append(ReviewFlag.MANUAL_REVIEW_REQUIRED)
        
        return flags
    
    def _flag_sheet_manual_review_items(self, column_mappings: List[ColumnMappingInfo],
                                       row_classifications: List[RowClassificationInfo],
                                       validation_summary: ValidationSummary) -> List[Dict[str, Any]]:
        """Flag items in a sheet that need manual review"""
        review_items = []
        
        # Low confidence column mappings
        for cm in column_mappings:
            if cm.confidence < self.review_thresholds['low_confidence']:
                review_items.append({
                    'type': 'low_confidence_column',
                    'column_index': cm.column_index,
                    'header': cm.original_header,
                    'confidence': cm.confidence,
                    'alternatives': cm.alternatives
                })
        
        # Rows with validation errors
        for rc in row_classifications:
            if rc.validation_errors:
                review_items.append({
                    'type': 'row_validation_errors',
                    'row_index': rc.row_index,
                    'errors': rc.validation_errors
                })
        
        # Rows with missing data
        for rc in row_classifications:
            if rc.completeness_score < self.review_thresholds['missing_data']:
                review_items.append({
                    'type': 'missing_data_row',
                    'row_index': rc.row_index,
                    'completeness_score': rc.completeness_score
                })
        
        return review_items
    
    def _generate_processing_notes(self, column_mapping: Any, 
                                 row_classification: Any,
                                 validation_result: Any) -> Tuple[List[str], List[str]]:
        """Generate processing notes and warnings"""
        notes = []
        warnings = []
        
        # Column mapping notes
        if hasattr(column_mapping, 'unmapped_columns'):
            if column_mapping.unmapped_columns:
                warnings.append(f"{len(column_mapping.unmapped_columns)} columns could not be mapped")
        elif column_mapping.get('unmapped_columns'):
            warnings.append(f"{len(column_mapping['unmapped_columns'])} columns could not be mapped")
        
        if hasattr(column_mapping, 'suggestions'):
            if column_mapping.suggestions:
                notes.extend(column_mapping.suggestions)
        elif column_mapping.get('suggestions'):
            notes.extend(column_mapping['suggestions'])
        
        # Row classification notes
        if hasattr(row_classification, 'suggestions'):
            if row_classification.suggestions:
                notes.extend(row_classification.suggestions)
        elif row_classification.get('suggestions'):
            notes.extend(row_classification['suggestions'])
        
        # Validation notes
        if hasattr(validation_result, 'suggestions'):
            if validation_result.suggestions:
                notes.extend(validation_result.suggestions)
        elif validation_result.get('suggestions'):
            notes.extend(validation_result['suggestions'])
        
        return notes, warnings
    
    def _identify_global_review_flags(self, sheet_mappings: List[SheetMapping]) -> List[ReviewFlag]:
        """Identify global review flags across all sheets"""
        flags = set()
        
        for sheet in sheet_mappings:
            flags.update(sheet.review_flags)
        
        return list(flags)
    
    def _is_export_ready(self, sheet_mappings: List[SheetMapping], global_confidence: float) -> bool:
        """Determine if the mapping is ready for export"""
        if global_confidence < 0.7:
            return False
        
        # Check if any sheets have critical issues
        for sheet in sheet_mappings:
            if sheet.processing_status == ProcessingStatus.FAILED:
                return False
            if ReviewFlag.VALIDATION_ERRORS in sheet.review_flags:
                return False
        
        return True
    
    def _generate_recommendations(self, sheet_mappings: List[SheetMapping]) -> List[str]:
        """Generate recommendations based on processing results"""
        recommendations = []
        
        # Check for low confidence sheets
        low_confidence_sheets = [s for s in sheet_mappings if s.overall_confidence < 0.6]
        if low_confidence_sheets:
            recommendations.append(f"Review {len(low_confidence_sheets)} sheets with low confidence")
        
        # Check for validation issues
        total_errors = sum(s.validation_summary.error_count for s in sheet_mappings)
        if total_errors > 0:
            recommendations.append(f"Fix {total_errors} validation errors across all sheets")
        
        # Check for missing data
        missing_data_sheets = [s for s in sheet_mappings if ReviewFlag.MISSING_DATA in s.review_flags]
        if missing_data_sheets:
            recommendations.append(f"Review data completeness in {len(missing_data_sheets)} sheets")
        
        # Check for ambiguous mappings
        ambiguous_sheets = [s for s in sheet_mappings if ReviewFlag.AMBIGUOUS_MAPPING in s.review_flags]
        if ambiguous_sheets:
            recommendations.append(f"Review ambiguous column mappings in {len(ambiguous_sheets)} sheets")
        
        return recommendations
    
    def _serialize_mapping(self, mapping: FileMapping) -> Dict[str, Any]:
        """Serialize mapping structure for JSON export"""
        mapping_dict = asdict(mapping)
        
        # Convert enums to strings
        mapping_dict['processing_summary'] = asdict(mapping.processing_summary)
        mapping_dict['metadata'] = asdict(mapping.metadata)
        
        for sheet in mapping_dict['sheets']:
            sheet['processing_status'] = sheet['processing_status'].value
            sheet['review_flags'] = [flag.value for flag in sheet['review_flags']]
            sheet['validation_summary'] = asdict(sheet['validation_summary'])
            
            for col_mapping in sheet['column_mappings']:
                col_mapping = asdict(col_mapping)
            
            for row_class in sheet['row_classifications']:
                row_class = asdict(row_class)
        
        mapping_dict['review_flags'] = [flag.value for flag in mapping.review_flags]
        
        return mapping_dict


# Convenience function for quick mapping generation
def generate_mapping_quick(processor_results: Dict[str, Any]) -> Dict[str, Any]:
    """
    Quick mapping generation
    
    Args:
        processor_results: Dictionary containing processing results
        
    Returns:
        Dictionary with mapping information
    """
    generator = MappingGenerator()
    mapping = generator.generate_file_mapping(processor_results)
    
    return {
        'global_confidence': mapping.global_confidence,
        'sheet_count': len(mapping.sheets),
        'export_ready': mapping.export_ready,
        'review_flags': [flag.value for flag in mapping.review_flags],
        'processing_summary': asdict(mapping.processing_summary)
    } 