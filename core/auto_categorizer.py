"""
Auto Categorizer for BOQ Tools
Automatically categorizes dataset rows using the CategoryDictionary
"""

import logging
import pandas as pd
from typing import Dict, List, Tuple, Optional, Any, Callable
from pathlib import Path
from dataclasses import dataclass, field

from core.category_dictionary import CategoryDictionary, CategoryMatch

logger = logging.getLogger(__name__)


@dataclass
class UnmatchedDescription:
    """Represents an unmatched description with metadata"""
    description: str
    source_sheet_name: str
    row_number: int
    original_index: int
    frequency: int = 1
    sample_rows: List[int] = field(default_factory=list)


@dataclass
class CategorizationResult:
    """Result of automatic categorization process"""
    dataframe: pd.DataFrame
    unmatched_descriptions: List[str]
    match_statistics: Dict[str, Any]
    total_rows: int
    matched_rows: int
    unmatched_rows: int
    match_rate: float


# Helper function to safely convert index to int
def safe_int(val):
    try:
        return int(val)
    except Exception:
        return 0


def collect_unmatched_descriptions(dataframe: pd.DataFrame,
                                 category_column: str = 'Category',
                                 description_column: str = 'Description',
                                 sheet_name_column: Optional[str] = None) -> List[UnmatchedDescription]:
    """
    Collect unmatched descriptions from categorized DataFrame
    
    Args:
        dataframe: Categorized DataFrame
        category_column: Name of the category column
        description_column: Name of the description column
        sheet_name_column: Name of the sheet name column (optional)
        
    Returns:
        List of UnmatchedDescription objects with metadata
    """
    logger.info(f"Collecting unmatched descriptions from {len(dataframe)} rows")
    
    # Validate inputs
    if category_column not in dataframe.columns:
        raise ValueError(f"Category column '{category_column}' not found in DataFrame")
    
    if description_column not in dataframe.columns:
        raise ValueError(f"Description column '{description_column}' not found in DataFrame")
    
    # Find rows with empty/null categories
    unmatched_mask = dataframe[category_column].isna() | (dataframe[category_column] == '') | (dataframe[category_column].isnull())
    unmatched_df = dataframe[unmatched_mask].copy()
    
    logger.info(f"Found {len(unmatched_df)} rows with unmatched descriptions")
    
    # Initialize collection for unique descriptions
    unique_descriptions: Dict[str, UnmatchedDescription] = {}
    
    # Process each unmatched row
    for index, row in unmatched_df.iterrows():
        description = str(row[description_column]).strip()
        
        # Skip empty descriptions
        if not description or description.lower() in ['nan', 'none', '']:
            continue
        
        # Normalize description for deduplication
        normalized_desc = description.lower()
        
        # Get source information
        source_sheet = str(row.get(sheet_name_column, 'Unknown')) if sheet_name_column else 'Unknown'
        row_number = safe_int(index) + 1  # Convert to 1-based row numbering
        
        if normalized_desc in unique_descriptions:
            # Update existing entry
            existing = unique_descriptions[normalized_desc]
            existing.frequency += 1
            existing.sample_rows.append(row_number)
            logger.debug(f"Duplicate description found: '{description[:50]}...' (frequency: {existing.frequency})")
        else:
            # Create new entry
            unmatched_desc = UnmatchedDescription(
                description=description,
                source_sheet_name=source_sheet,
                row_number=row_number,
                original_index=safe_int(index),
                frequency=1,
                sample_rows=[row_number]
            )
            unique_descriptions[normalized_desc] = unmatched_desc
            logger.debug(f"New unmatched description: '{description[:50]}...'")
    
    # Convert to list and sort by frequency (most frequent first)
    unmatched_list = list(unique_descriptions.values())
    unmatched_list.sort(key=lambda x: x.frequency, reverse=True)
    
    logger.info(f"Collected {len(unmatched_list)} unique unmatched descriptions")
    logger.info(f"Total frequency: {sum(desc.frequency for desc in unmatched_list)}")
    
    # Log some statistics
    if unmatched_list:
        max_freq = max(desc.frequency for desc in unmatched_list)
        min_freq = min(desc.frequency for desc in unmatched_list)
        avg_freq = sum(desc.frequency for desc in unmatched_list) / len(unmatched_list)
        
        logger.info(f"Frequency statistics:")
        logger.info(f"  Max frequency: {max_freq}")
        logger.info(f"  Min frequency: {min_freq}")
        logger.info(f"  Average frequency: {avg_freq:.2f}")
        
        # Show top unmatched descriptions
        logger.info(f"Top 5 unmatched descriptions:")
        for i, desc in enumerate(unmatched_list[:5]):
            logger.info(f"  {i+1}. '{desc.description[:60]}...' (frequency: {desc.frequency})")
    
    return unmatched_list


def auto_categorize_dataset(dataframe: pd.DataFrame, 
                          category_dictionary: CategoryDictionary,
                          description_column: str = 'Description',
                          category_column: str = 'Category',
                          confidence_threshold: float = 0.8,
                          progress_callback: Optional[Callable] = None) -> CategorizationResult:
    """
    Automatically categorize dataset rows using the CategoryDictionary
    
    Args:
        dataframe: Pandas DataFrame that has gone through mapping process
        category_dictionary: CategoryDictionary instance
        description_column: Name of the column containing descriptions
        category_column: Name of the column to add for categories
        confidence_threshold: Minimum confidence threshold for matches
        progress_callback: Optional callback function for progress tracking
        
    Returns:
        CategorizationResult with categorized DataFrame and statistics
    """
    logger.info(f"Starting automatic categorization of {len(dataframe)} rows")
    
    # Validate inputs
    if description_column not in dataframe.columns:
        raise ValueError(f"Description column '{description_column}' not found in DataFrame")
    
    # Create a copy to avoid modifying the original
    df = dataframe.copy()
    
    # Initialize results tracking
    total_rows = len(df)
    matched_rows = 0
    unmatched_rows = 0
    unmatched_descriptions = []
    match_types = {'exact': 0, 'partial': 0, 'fuzzy': 0, 'none': 0}
    
    # Add category column if it doesn't exist
    if category_column not in df.columns:
        df[category_column] = ''
    
    # Process each row
    for index, row in df.iterrows():
        try:
            description = str(row[description_column]).strip()
            if not description or description.lower() in ['nan', 'none', '']:
                unmatched_rows += 1
                continue

            match = category_dictionary.find_category(description, confidence_threshold)
            match_types[match.match_type] += 1

            if match.matched_category:
                df.at[index, category_column] = match.matched_category
                matched_rows += 1
                if match.match_type == 'fuzzy':
                    # Find the best matching dictionary string for fuzzy
                    best_dict_str = None
                    best_similarity = 0.0
                    for dict_desc, mapping in category_dictionary.mappings.items():
                        similarity = category_dictionary._calculate_fuzzy_similarity(description.lower().strip(), dict_desc)
                        if similarity > best_similarity:
                            best_similarity = similarity
                            best_dict_str = dict_desc
                    print(f"[CATEGORIZATION] '{description}' → '{match.matched_category}' (type: {match.match_type}, confidence: {match.confidence:.2f}, matched to: '{best_dict_str}')")
                else:
                    print(f"[CATEGORIZATION] '{description}' → '{match.matched_category}' (type: {match.match_type}, confidence: {match.confidence:.2f})")
            else:
                df.at[index, category_column] = ''
                unmatched_rows += 1
                print(f"[CATEGORIZATION] '{description}' → UNMATCHED")
            
            # Progress callback
            if progress_callback:
                progress = (index + 1) / total_rows * 100
                progress_callback(progress, f"Processing row {index + 1}/{total_rows}")
            
        except Exception as e:
            logger.error(f"Error processing row {index}: {e}")
            unmatched_rows += 1
            unmatched_descriptions.append(f"ERROR: {str(row.get(description_column, 'Unknown'))}")
            continue
    
    # Calculate final statistics
    match_rate = matched_rows / total_rows if total_rows > 0 else 0.0
    
    match_statistics = {
        'total_rows': total_rows,
        'matched_rows': matched_rows,
        'unmatched_rows': unmatched_rows,
        'match_rate': match_rate,
        'match_types': match_types,
        'confidence_threshold': confidence_threshold,
        'unique_categories_found': len(df[category_column].dropna().unique()),
        'category_distribution': df[category_column].value_counts().to_dict()
    }
    
    # Log results
    logger.info(f"Categorization completed:")
    logger.info(f"  Total rows: {total_rows}")
    logger.info(f"  Matched rows: {matched_rows} ({match_rate:.1%})")
    logger.info(f"  Unmatched rows: {unmatched_rows}")
    logger.info(f"  Match types: {match_types}")
    logger.info(f"  Unique categories assigned: {match_statistics['unique_categories_found']}")
    
    # Create result object
    result = CategorizationResult(
        dataframe=df,
        unmatched_descriptions=unmatched_descriptions,
        match_statistics=match_statistics,
        total_rows=total_rows,
        matched_rows=matched_rows,
        unmatched_rows=unmatched_rows,
        match_rate=match_rate
    )
    
    return result


class AutoCategorizer:
    """
    Automatic dataset categorizer using CategoryDictionary
    """
    
    def __init__(self, category_dictionary: CategoryDictionary):
        """
        Initialize the auto categorizer
        
        Args:
            category_dictionary: CategoryDictionary instance to use for categorization
        """
        self.category_dictionary = category_dictionary
        logger.info("Auto Categorizer initialized")
    
    def auto_categorize_dataset(self, dataframe: pd.DataFrame, 
                               description_column: str = 'Description',
                               category_column: str = 'Category',
                               confidence_threshold: float = 0.8,
                               progress_callback: Optional[Callable] = None) -> CategorizationResult:
        """
        Automatically categorize dataset rows using the CategoryDictionary
        
        Args:
            dataframe: Pandas DataFrame that has gone through mapping process
            description_column: Name of the column containing descriptions
            category_column: Name of the column to add for categories
            confidence_threshold: Minimum confidence threshold for matches
            progress_callback: Optional callback function for progress tracking
            
        Returns:
            CategorizationResult with categorized DataFrame and statistics
        """
        return auto_categorize_dataset(
            dataframe, self.category_dictionary, description_column, 
            category_column, confidence_threshold, progress_callback
        )
    
    def collect_unmatched_descriptions(self, dataframe: pd.DataFrame,
                                     category_column: str = 'Category',
                                     description_column: str = 'Description',
                                     sheet_name_column: Optional[str] = None) -> List[UnmatchedDescription]:
        """
        Collect unmatched descriptions from categorized DataFrame
        
        Args:
            dataframe: Categorized DataFrame
            category_column: Name of the category column
            description_column: Name of the description column
            sheet_name_column: Name of the sheet name column (optional)
            
        Returns:
            List of UnmatchedDescription objects with metadata
        """
        return collect_unmatched_descriptions(
            dataframe, category_column, description_column, sheet_name_column
        )
    
    def categorize_with_learning(self, dataframe: pd.DataFrame,
                                description_column: str = 'Description',
                                category_column: str = 'Category',
                                confidence_threshold: float = 0.8,
                                auto_learn_threshold: float = 0.9,
                                progress_callback: Optional[Callable] = None) -> CategorizationResult:
        """
        Categorize dataset with automatic learning of high-confidence matches
        
        Args:
            dataframe: Pandas DataFrame to categorize
            description_column: Name of the description column
            category_column: Name of the category column
            confidence_threshold: Minimum confidence for categorization
            auto_learn_threshold: Minimum confidence for automatic learning
            progress_callback: Optional progress callback
            
        Returns:
            CategorizationResult with categorized DataFrame
        """
        logger.info(f"Starting categorization with learning for {len(dataframe)} rows")
        
        # First, categorize normally
        result = self.auto_categorize_dataset(
            dataframe, description_column, category_column, 
            confidence_threshold, progress_callback
        )
        
        # Learn from high-confidence matches
        learned_count = 0
        for index, row in result.dataframe.iterrows():
            description = str(row[description_column]).strip()
            category = str(row[category_column]).strip()
            
            if description and category:
                # Check if this is a high-confidence match that should be learned
                match = self.category_dictionary.find_category(description, auto_learn_threshold)
                
                if match.confidence >= auto_learn_threshold and match.matched_category == category:
                    # This is a high-confidence match, add to dictionary if not already present
                    if description.lower() not in self.category_dictionary.mappings:
                        success = self.category_dictionary.add_mapping(
                            description, category, 
                            confidence=match.confidence,
                            notes="Auto-learned from dataset categorization"
                        )
                        if success:
                            learned_count += 1
        
        if learned_count > 0:
            logger.info(f"Auto-learned {learned_count} new mappings")
            # Save the updated dictionary
            self.category_dictionary.save_dictionary()
        
        return result
    
    def get_categorization_summary(self, result: CategorizationResult) -> str:
        """
        Generate a summary report of the categorization results
        
        Args:
            result: CategorizationResult from auto_categorize_dataset
            
        Returns:
            Formatted summary string
        """
        stats = result.match_statistics
        
        summary = f"""
Categorization Summary
=====================
Total Rows Processed: {stats['total_rows']}
Matched Rows: {stats['matched_rows']} ({stats['match_rate']:.1%})
Unmatched Rows: {stats['unmatched_rows']}

Match Type Breakdown:
  - Exact Matches: {stats['match_types']['exact']}
  - Partial Matches: {stats['match_types']['partial']}
  - Fuzzy Matches: {stats['match_types']['fuzzy']}
  - No Matches: {stats['match_types']['none']}

Categories Found: {stats['unique_categories_found']}
Confidence Threshold: {stats['confidence_threshold']}

Top Categories:
"""
        
        # Add top categories
        category_dist = stats['category_distribution']
        if category_dist:
            sorted_categories = sorted(category_dist.items(), key=lambda x: x[1], reverse=True)
            for category, count in sorted_categories[:10]:
                if category:  # Skip empty categories
                    summary += f"  - {category}: {count} items\n"
        
        if result.unmatched_descriptions:
            summary += f"\nUnmatched Descriptions: {len(result.unmatched_descriptions)}"
            summary += f"\nSample unmatched descriptions:\n"
            for desc in result.unmatched_descriptions[:5]:
                summary += f"  - {desc[:60]}...\n"
        
        return summary
    
    def export_categorization_report(self, result: CategorizationResult, 
                                   output_path: Path) -> bool:
        """
        Export categorization results to a detailed report
        
        Args:
            result: CategorizationResult from auto_categorize_dataset
            output_path: Path to save the report
            
        Returns:
            True if export was successful
        """
        try:
            # Create detailed report
            report = {
                'summary': result.match_statistics,
                'unmatched_descriptions': result.unmatched_descriptions,
                'category_distribution': result.match_statistics['category_distribution'],
                'match_type_breakdown': result.match_statistics['match_types']
            }
            
            # Save as JSON
            import json
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Categorization report exported to {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error exporting categorization report: {e}")
            return False 