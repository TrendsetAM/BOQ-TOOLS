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
    
    # Processing results
    column_mappings: List[ColumnMappingInfo]
    row_classifications: List[RowClassificationInfo]
    validation_summary: ValidationSummary
    
    # Confidence scores
    overall_confidence: float
    column_mapping_confidence: float
    row_classification_confidence: float
    data_quality_confidence: float
    
    # Review flags
    review_flags: List[ReviewFlag]
    manual_review_items: List[Dict[str, Any]]
    
    # Processing metadata
    processing_notes: List[str]
    warnings: List[str]
    processing_time: float


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
            
            # Create metadata
            metadata = self._create_file_metadata(file_info)
            
            # Process each sheet
            sheet_mappings = []
            for sheet_name in sheet_data.keys():
                sheet_mapping = self._create_sheet_mapping(
                    sheet_name=sheet_name,
                    sheet_data=sheet_data.get(sheet_name, []),
                    column_mapping=column_mappings.get(sheet_name, {}),
                    row_classification=row_classifications.get(sheet_name, {}),
                    validation_result=validation_results.get(sheet_name, {})
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
                             column_mapping: Dict[str, Any], row_classification: Dict[str, Any],
                             validation_result: Dict[str, Any]) -> SheetMapping:
        """Create detailed sheet mapping"""
        
        # Extract column mapping information
        column_mappings = []
        for col_info in column_mapping.get('mappings', []):
            col_mapping = ColumnMappingInfo(
                column_index=col_info.get('column_index', 0),
                original_header=col_info.get('original_header', ''),
                normalized_header=col_info.get('normalized_header', ''),
                mapped_type=col_info.get('mapped_type', ''),
                confidence=col_info.get('confidence', 0.0),
                alternatives=col_info.get('alternatives', []),
                reasoning=col_info.get('reasoning', []),
                is_required=col_info.get('is_required', False),
                validation_status=col_info.get('validation_status', 'unknown')
            )
            column_mappings.append(col_mapping)
        
        # Extract row classification information
        row_classifications = []
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
        validation_summary = ValidationSummary(
            overall_score=validation_result.get('overall_score', 0.0),
            mathematical_consistency=validation_result.get('mathematical_consistency', 0.0),
            data_type_quality=validation_result.get('data_type_quality', 0.0),
            business_rule_compliance=validation_result.get('business_rule_compliance', 0.0),
            error_count=validation_result.get('error_count', 0),
            warning_count=validation_result.get('warning_count', 0),
            info_count=validation_result.get('info_count', 0),
            suggestions=validation_result.get('suggestions', [])
        )
        
        # Calculate confidence scores
        column_mapping_confidence = column_mapping.get('overall_confidence', 0.0)
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
        
        return SheetMapping(
            sheet_name=sheet_name,
            processing_status=processing_status,
            row_count=len(sheet_data),
            column_count=len(sheet_data[0]) if sheet_data else 0,
            header_row_index=column_mapping.get('header_row_index', 0),
            header_confidence=column_mapping.get('header_confidence', 0.0),
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
            processing_time=0.0  # Would be calculated during actual processing
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
    
    def _generate_processing_notes(self, column_mapping: Dict[str, Any], 
                                 row_classification: Dict[str, Any],
                                 validation_result: Dict[str, Any]) -> Tuple[List[str], List[str]]:
        """Generate processing notes and warnings"""
        notes = []
        warnings = []
        
        # Column mapping notes
        if column_mapping.get('unmapped_columns'):
            warnings.append(f"{len(column_mapping['unmapped_columns'])} columns could not be mapped")
        
        if column_mapping.get('suggestions'):
            notes.extend(column_mapping['suggestions'])
        
        # Row classification notes
        if row_classification.get('suggestions'):
            notes.extend(row_classification['suggestions'])
        
        # Validation notes
        if validation_result.get('suggestions'):
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