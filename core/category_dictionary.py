"""
Category Dictionary for BOQ Tools
Manages predefined dictionary mapping descriptions to categories for automatic row categorization
"""

import json
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Set, Any
from dataclasses import dataclass, asdict
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
    suggestions: List[str] = None


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
        self.dictionary_file = dictionary_file or Path("config/category_dictionary.json")
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
                logger.info(f"Dictionary file not found, creating new: {self.dictionary_file}")
                self._create_default_dictionary()
                
        except Exception as e:
            logger.error(f"Error loading dictionary: {e}")
            self._create_default_dictionary()
    
    def _create_default_dictionary(self) -> None:
        """Create default category dictionary with common BOQ categories"""
        default_mappings = [
            CategoryMapping("concrete", CategoryType.MATERIALS.value),
            CategoryMapping("steel", CategoryType.MATERIALS.value),
            CategoryMapping("cement", CategoryType.MATERIALS.value),
            CategoryMapping("sand", CategoryType.MATERIALS.value),
            CategoryMapping("aggregate", CategoryType.MATERIALS.value),
            CategoryMapping("brick", CategoryType.MATERIALS.value),
            CategoryMapping("block", CategoryType.MATERIALS.value),
            CategoryMapping("timber", CategoryType.MATERIALS.value),
            CategoryMapping("paint", CategoryType.MATERIALS.value),
            CategoryMapping("tiles", CategoryType.MATERIALS.value),
            CategoryMapping("glass", CategoryType.MATERIALS.value),
            CategoryMapping("insulation", CategoryType.MATERIALS.value),
            CategoryMapping("roofing", CategoryType.MATERIALS.value),
            CategoryMapping("excavation", CategoryType.LABOR.value),
            CategoryMapping("foundation", CategoryType.LABOR.value),
            CategoryMapping("carpentry", CategoryType.LABOR.value),
            CategoryMapping("electrical", CategoryType.LABOR.value),
            CategoryMapping("plumbing", CategoryType.LABOR.value),
            CategoryMapping("masonry", CategoryType.LABOR.value),
            CategoryMapping("painting", CategoryType.LABOR.value),
            CategoryMapping("finishing", CategoryType.LABOR.value),
            CategoryMapping("crane", CategoryType.EQUIPMENT.value),
            CategoryMapping("excavator", CategoryType.EQUIPMENT.value),
            CategoryMapping("bulldozer", CategoryType.EQUIPMENT.value),
            CategoryMapping("loader", CategoryType.EQUIPMENT.value),
            CategoryMapping("compactor", CategoryType.EQUIPMENT.value),
            CategoryMapping("generator", CategoryType.EQUIPMENT.value),
            CategoryMapping("welding", CategoryType.EQUIPMENT.value),
            CategoryMapping("testing", CategoryType.SERVICES.value),
            CategoryMapping("inspection", CategoryType.SERVICES.value),
            CategoryMapping("certification", CategoryType.SERVICES.value),
            CategoryMapping("permits", CategoryType.SERVICES.value),
            CategoryMapping("design", CategoryType.SERVICES.value),
            CategoryMapping("consulting", CategoryType.SERVICES.value),
            CategoryMapping("overhead", CategoryType.OVERHEAD.value),
            CategoryMapping("site office", CategoryType.OVERHEAD.value),
            CategoryMapping("utilities", CategoryType.OVERHEAD.value),
            CategoryMapping("insurance", CategoryType.OVERHEAD.value),
            CategoryMapping("profit", CategoryType.PROFIT.value),
            CategoryMapping("margin", CategoryType.PROFIT.value),
            CategoryMapping("contingency", CategoryType.CONTINGENCY.value),
            CategoryMapping("allowance", CategoryType.CONTINGENCY.value),
            CategoryMapping("variation", CategoryType.CONTINGENCY.value),
            CategoryMapping("tax", CategoryType.TAXES.value),
            CategoryMapping("vat", CategoryType.TAXES.value),
            CategoryMapping("duty", CategoryType.TAXES.value),
        ]
        
        for mapping in default_mappings:
            self.mappings[mapping.description.lower()] = mapping
            self.categories.add(mapping.category)
        
        # Add all category types
        for category_type in CategoryType:
            self.categories.add(category_type.value)
        
        logger.info("Created default category dictionary")
    
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
            category: Category to assign
            confidence: Confidence level (0.0 to 1.0)
            notes: Optional notes about the mapping
            
        Returns:
            True if mapping was added successfully
        """
        try:
            # Normalize description
            normalized_desc = description.lower().strip()
            
            if not normalized_desc:
                logger.warning("Cannot add mapping with empty description")
                return False
            
            # Create new mapping
            mapping = CategoryMapping(
                description=normalized_desc,
                category=category.lower().strip(),
                confidence=confidence,
                notes=notes
            )
            
            # Add to dictionary
            self.mappings[normalized_desc] = mapping
            self.categories.add(category.lower().strip())
            
            logger.info(f"Added mapping: '{description}' -> '{category}'")
            return True
            
        except Exception as e:
            logger.error(f"Error adding mapping: {e}")
            return False
    
    def find_category(self, description: str, threshold: float = 0.8) -> CategoryMatch:
        """
        Find category for a given description
        
        Args:
            description: Item description to categorize
            threshold: Minimum confidence threshold for matches
            
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
        
        # Try exact match first
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
        
        # Try partial matches
        partial_matches = []
        for dict_desc, mapping in self.mappings.items():
            if self._is_partial_match(normalized_desc, dict_desc):
                partial_matches.append((mapping, self._calculate_partial_confidence(normalized_desc, dict_desc)))
        
        if partial_matches:
            # Sort by confidence and return best match
            partial_matches.sort(key=lambda x: x[1], reverse=True)
            best_match, confidence = partial_matches[0]
            
            if confidence >= threshold:
                best_match.usage_count += 1
                return CategoryMatch(
                    description=normalized_desc,
                    matched_category=best_match.category,
                    confidence=confidence,
                    match_type='partial',
                    original_description=original_desc,
                    suggestions=[m.category for m, _ in partial_matches[1:3]]  # Top 3 suggestions
                )
        
        # Try fuzzy matching
        fuzzy_matches = []
        for dict_desc, mapping in self.mappings.items():
            similarity = self._calculate_fuzzy_similarity(normalized_desc, dict_desc)
            if similarity >= threshold:
                fuzzy_matches.append((mapping, similarity))
        
        if fuzzy_matches:
            fuzzy_matches.sort(key=lambda x: x[1], reverse=True)
            best_match, confidence = fuzzy_matches[0]
            
            best_match.usage_count += 1
            return CategoryMatch(
                description=normalized_desc,
                matched_category=best_match.category,
                confidence=confidence,
                match_type='fuzzy',
                original_description=original_desc,
                suggestions=[m.category for m, _ in fuzzy_matches[1:3]]
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
    
    def _is_partial_match(self, desc1: str, desc2: str) -> bool:
        """Check if two descriptions partially match"""
        words1 = set(desc1.split())
        words2 = set(desc2.split())
        
        if not words1 or not words2:
            return False
        
        # Check if any significant words match
        common_words = words1.intersection(words2)
        significant_words = [w for w in common_words if len(w) > 2]  # Ignore short words
        
        return len(significant_words) > 0
    
    def _calculate_partial_confidence(self, desc1: str, desc2: str) -> float:
        """Calculate confidence for partial matches"""
        words1 = set(desc1.split())
        words2 = set(desc2.split())
        
        if not words1 or not words2:
            return 0.0
        
        common_words = words1.intersection(words2)
        significant_words = [w for w in common_words if len(w) > 2]
        
        # Calculate Jaccard similarity for significant words
        if len(significant_words) == 0:
            return 0.0
        
        union_size = len(words1.union(words2))
        if union_size == 0:
            return 0.0
        
        return len(significant_words) / union_size
    
    def _calculate_fuzzy_similarity(self, desc1: str, desc2: str) -> float:
        """Calculate fuzzy string similarity using simple algorithm"""
        if desc1 == desc2:
            return 1.0
        
        if not desc1 or not desc2:
            return 0.0
        
        # Simple character-based similarity
        len1, len2 = len(desc1), len(desc2)
        max_len = max(len1, len2)
        
        if max_len == 0:
            return 0.0
        
        # Count matching characters in sequence
        matches = 0
        i, j = 0, 0
        
        while i < len1 and j < len2:
            if desc1[i] == desc2[j]:
                matches += 1
                i += 1
                j += 1
            elif len1 > len2:
                i += 1
            else:
                j += 1
        
        return matches / max_len
    
    def get_all_categories(self) -> List[str]:
        """Get all available categories"""
        return sorted(list(self.categories))
    
    def get_mappings_for_category(self, category: str) -> List[CategoryMapping]:
        """Get all mappings for a specific category"""
        category_lower = category.lower()
        return [mapping for mapping in self.mappings.values() 
                if mapping.category == category_lower]
    
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
            mapping.category = new_category.lower().strip()
            
            if new_confidence is not None:
                mapping.confidence = new_confidence
            
            if new_notes is not None:
                mapping.notes = new_notes
            
            self.categories.add(new_category.lower().strip())
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