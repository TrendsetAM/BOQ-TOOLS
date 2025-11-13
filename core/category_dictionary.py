"""
Category Dictionary for BOQ Tools
Manages predefined dictionary mapping descriptions to categories for automatic row categorization
"""

import json
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Any
from dataclasses import dataclass, asdict, field
from enum import Enum

logger = logging.getLogger(__name__)


class CategoryType(Enum):
    """Types of categories for BOQ items"""
    MATERIALS = "materials"
    LABOR = "labor"
    EQUIPMENT = "equipment"
    SERVICES = "services"
    OVERHEAD = "overhead"
    PROFIT = "profit"
    CONTINGENCY = "contingency"
    TAXES = "taxes"
    OTHER = "other"


@dataclass
class CategoryMapping:
    """Individual mapping from description to category"""
    description: str
    category: str
    confidence: float = 1.0
    created_date: Optional[str] = None
    usage_count: int = 0
    notes: Optional[str] = None


@dataclass
class CategoryMatch:
    """Result of category matching"""
    description: str
    matched_category: Optional[str]
    confidence: float
    match_type: str  # 'exact', 'partial', 'fuzzy', 'none'
    original_description: str
    suggestions: List[str] = field(default_factory=list)


class CategoryDictionary:
    """
    Manages a dictionary mapping descriptions to categories for automatic row categorization
    """
    
    def __init__(self, dictionary_file: Optional[Path] = None):
        """
        Initialize the category dictionary
        
        Args:
            dictionary_file: Path to JSON file containing the category dictionary
        """
        if dictionary_file is None:
            # Use the user config directory for standalone executable compatibility
            from utils.config import get_user_config_path
            user_config_file = get_user_config_path('category_dictionary.json')
            self.dictionary_file = Path(user_config_file)
        else:
            self.dictionary_file = dictionary_file
        self.mappings: Dict[str, CategoryMapping] = {}
        self.categories: Set[str] = set()
        self._load_dictionary()
        
        logger.info(f"Category Dictionary initialized with {len(self.mappings)} mappings")
    
    def _load_dictionary(self) -> None:
        """Load dictionary from JSON file"""
        try:
            if self.dictionary_file.exists():
                with open(self.dictionary_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Load mappings
                if 'mappings' in data:
                    for mapping_data in data['mappings']:
                        mapping = CategoryMapping(**mapping_data)
                        self.mappings[mapping.description.lower()] = mapping
                        self.categories.add(mapping.category)
                
                # Load categories if provided
                if 'categories' in data:
                    self.categories.update(data['categories'])
                
                logger.info(f"Loaded {len(self.mappings)} mappings from {self.dictionary_file}")
            else:
                logger.info(f"Dictionary file not found at {self.dictionary_file}")
                # Try to copy from bundled default if running as executable
                self._try_copy_bundled_dictionary()
                
                # If still no file, create minimal default
                if not self.dictionary_file.exists():
                    logger.info("Creating minimal default dictionary")
                    self._create_default_dictionary()
                    self.save_dictionary()  # Save the default to user directory
                else:
                    # Reload after copying bundled version
                    self._load_dictionary()
                
        except Exception as e:
            logger.error(f"Error loading dictionary: {e}")
            self._create_default_dictionary()
            self.save_dictionary()  # Save the default to user directory
    
    def _create_default_dictionary(self) -> None:
        """Create default category dictionary with common BOQ categories"""
        # Import the actual categories used in the system
        from core.manual_categorizer import get_manual_categorization_categories
        
        # Get the pretty categories
        pretty_categories = get_manual_categorization_categories()
        
        # Create some basic default mappings using pretty categories
        default_mappings = [
            CategoryMapping("concrete", "Civil Works"),
            CategoryMapping("steel", "Civil Works"),
            CategoryMapping("cement", "Civil Works"),
            CategoryMapping("sand", "Civil Works"),
            CategoryMapping("aggregate", "Civil Works"),
            CategoryMapping("excavation", "Earth Movement"),
            CategoryMapping("foundation", "Civil Works"),
            CategoryMapping("electrical", "Electrical Works"),
            CategoryMapping("cable", "Electrical Works"),
            CategoryMapping("solar", "PV Mod. Installation"),
            CategoryMapping("panel", "PV Mod. Installation"),
            CategoryMapping("inverter", "Electrical Works"),
            CategoryMapping("transformer", "Electrical Works"),
            CategoryMapping("trenching", "Trenching"),
            CategoryMapping("road", "Roads"),
            CategoryMapping("access", "Roads"),
            CategoryMapping("building", "OEM Building"),
            CategoryMapping("office", "OEM Building"),
            CategoryMapping("tracker", "Tracker Inst."),
            CategoryMapping("mounting", "Tracker Inst."),
            CategoryMapping("cleaning", "Cleaning and Cabling of PV Mod."),
            CategoryMapping("cabling", "Cleaning and Cabling of PV Mod."),
            CategoryMapping("overhead", "General Costs"),
            CategoryMapping("management", "General Costs"),
            CategoryMapping("permit", "General Costs"),
            CategoryMapping("insurance", "General Costs"),
            CategoryMapping("other", "Other"),
        ]
        
        for mapping in default_mappings:
            self.mappings[mapping.description.lower()] = mapping
            self.categories.add(mapping.category)
        
        # Add all pretty categories to ensure they're available
        for category in pretty_categories:
            self.categories.add(category)
        
        logger.info("Created default category dictionary with pretty categories")
    
    def _try_copy_bundled_dictionary(self) -> None:
        """Try to copy bundled dictionary from executable to user directory"""
        try:
            import os
            # Check if running as executable with bundled config
            if 'BOQ_TOOLS_BUNDLE_DIR' in os.environ:
                bundle_dict_path = Path(os.environ['BOQ_TOOLS_BUNDLE_DIR']) / 'config' / 'category_dictionary.json'
                if bundle_dict_path.exists():
                    # Ensure user config directory exists
                    self.dictionary_file.parent.mkdir(parents=True, exist_ok=True)
                    # Copy the bundled dictionary to user directory
                    import shutil
                    shutil.copy2(bundle_dict_path, self.dictionary_file)
                    logger.info(f"Copied bundled dictionary from {bundle_dict_path} to {self.dictionary_file}")
                else:
                    logger.debug(f"No bundled dictionary found at {bundle_dict_path}")
            else:
                logger.debug("Not running as executable, skipping bundled dictionary copy")
        except Exception as e:
            logger.warning(f"Could not copy bundled dictionary: {e}")
    
    def save_dictionary(self) -> bool:
        """Save dictionary to JSON file"""
        try:
            # Ensure directory exists
            self.dictionary_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Prepare data for saving
            data = {
                'mappings': [asdict(mapping) for mapping in self.mappings.values()],
                'categories': list(self.categories),
                'metadata': {
                    'total_mappings': len(self.mappings),
                    'total_categories': len(self.categories),
                    'last_updated': str(Path().cwd())
                }
            }
            
            with open(self.dictionary_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Dictionary saved to {self.dictionary_file}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving dictionary: {e}")
            return False
    
    def add_mapping(self, description: str, category: str, confidence: float = 1.0, 
                   notes: Optional[str] = None) -> bool:
        """
        Add new description-category mapping
        
        Args:
            description: Item description
            category: Category to assign (in pretty format)
            confidence: Confidence level (0.0 to 1.0)
            notes: Optional notes about the mapping
            
        Returns:
            True if mapping was added successfully
        """
        try:
            # Normalize description (keep lowercase for matching)
            normalized_desc = description.lower().strip()
            
            if not normalized_desc:
                logger.warning("Cannot add mapping with empty description")
                return False
            
            # Keep category in pretty format (no lowercasing)
            pretty_category = category.strip()
            
            # Create new mapping
            mapping = CategoryMapping(
                description=normalized_desc,
                category=pretty_category,  # Store in pretty format
                confidence=confidence,
                notes=notes
            )
            
            # Add to dictionary
            self.mappings[normalized_desc] = mapping
            self.categories.add(pretty_category)  # Store pretty category
            
            logger.info(f"Added mapping: '{description}' -> '{pretty_category}'")
            return True
            
        except Exception as e:
            logger.error(f"Error adding mapping: {e}")
            return False
    
    def find_category(self, description: str, threshold: float = 0.8) -> CategoryMatch:
        """
        Find category for a given description (exact matches only)
        
        Args:
            description: Item description to categorize
            threshold: Minimum confidence threshold for matches (not used for exact matches)
            
        Returns:
            CategoryMatch with match results
        """
        original_desc = description
        normalized_desc = description.lower().strip()
        
        if not normalized_desc:
            return CategoryMatch(
                description=normalized_desc,
                matched_category=None,
                confidence=0.0,
                match_type='none',
                original_description=original_desc,
                suggestions=[]
            )
        
        # Try exact match only
        if normalized_desc in self.mappings:
            mapping = self.mappings[normalized_desc]
            mapping.usage_count += 1
            return CategoryMatch(
                description=normalized_desc,
                matched_category=mapping.category,
                confidence=mapping.confidence,
                match_type='exact',
                original_description=original_desc,
                suggestions=[]
            )
        
        # No match found
        return CategoryMatch(
            description=normalized_desc,
            matched_category=None,
            confidence=0.0,
            match_type='none',
            original_description=original_desc,
            suggestions=list(self.categories)[:5]  # Top 5 categories as suggestions
        )
    
    def get_all_categories(self) -> List[str]:
        """Get all available categories"""
        return sorted(list(self.categories))
    
    def get_mappings_for_category(self, category: str) -> List[CategoryMapping]:
        """Get all mappings for a specific category"""
        # Compare with pretty category format
        return [mapping for mapping in self.mappings.values() 
                if mapping.category == category.strip()]
    
    def remove_mapping(self, description: str) -> bool:
        """Remove a mapping from the dictionary"""
        normalized_desc = description.lower().strip()
        
        if normalized_desc in self.mappings:
            del self.mappings[normalized_desc]
            logger.info(f"Removed mapping: '{description}'")
            return True
        
        return False
    
    def update_mapping(self, description: str, new_category: str, 
                      new_confidence: Optional[float] = None,
                      new_notes: Optional[str] = None) -> bool:
        """Update an existing mapping"""
        normalized_desc = description.lower().strip()
        
        if normalized_desc in self.mappings:
            mapping = self.mappings[normalized_desc]
            mapping.category = new_category.strip()  # Store in pretty format
            
            if new_confidence is not None:
                mapping.confidence = new_confidence
            
            if new_notes is not None:
                mapping.notes = new_notes
            
            self.categories.add(new_category.strip())  # Store pretty category
            logger.info(f"Updated mapping: '{description}' -> '{new_category}'")
            return True
        
        return False
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get dictionary statistics"""
        category_counts = {}
        for mapping in self.mappings.values():
            category_counts[mapping.category] = category_counts.get(mapping.category, 0) + 1
        
        return {
            'total_mappings': len(self.mappings),
            'total_categories': len(self.categories),
            'categories': list(self.categories),
            'category_counts': category_counts,
            'most_used_mappings': sorted(
                self.mappings.values(), 
                key=lambda x: x.usage_count, 
                reverse=True
            )[:10]
        }
    
    def export_dictionary(self, export_path: Path) -> bool:
        """Export dictionary to a different file"""
        try:
            data = {
                'mappings': [asdict(mapping) for mapping in self.mappings.values()],
                'categories': list(self.categories),
                'statistics': self.get_statistics()
            }
            
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Dictionary exported to {export_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error exporting dictionary: {e}")
            return False
    
    def import_dictionary(self, import_path: Path, merge: bool = True) -> bool:
        """Import dictionary from a file"""
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            imported_count = 0
            
            if 'mappings' in data:
                for mapping_data in data['mappings']:
                    mapping = CategoryMapping(**mapping_data)
                    normalized_desc = mapping.description.lower()
                    
                    if merge or normalized_desc not in self.mappings:
                        self.mappings[normalized_desc] = mapping
                        self.categories.add(mapping.category)
                        imported_count += 1
            
            logger.info(f"Imported {imported_count} mappings from {import_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error importing dictionary: {e}")
            return False 