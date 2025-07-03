#!/usr/bin/env python3
"""
Migration script to convert category dictionary from lowercase internal categories to pretty categories.
This script should be run once after updating the system to use pretty categories only.
"""

import json
from pathlib import Path
from core.manual_categorizer import get_manual_categorization_categories

def migrate_category_dictionary():
    """Migrate existing category dictionary to use pretty categories"""
    
    # Get the mapping from lowercase to pretty
    pretty_categories = get_manual_categorization_categories()
    lowercase_to_pretty = {cat.lower(): cat for cat in pretty_categories}
    
    print("Pretty categories mapping:")
    for lower, pretty in lowercase_to_pretty.items():
        print(f"  '{lower}' -> '{pretty}'")
    
    # Load existing dictionary
    dict_path = Path('config/category_dictionary.json')
    if not dict_path.exists():
        print("No existing dictionary found - nothing to migrate")
        return
    
    print(f"\nLoading existing dictionary from {dict_path}")
    with open(dict_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Create backup
    backup_path = dict_path.with_suffix('.backup.json')
    print(f"Creating backup at {backup_path}")
    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    # Migrate categories
    old_categories = data.get('categories', [])
    new_categories = []
    
    print(f"\nMigrating {len(old_categories)} categories:")
    for old_cat in old_categories:
        if old_cat.lower() in lowercase_to_pretty:
            new_cat = lowercase_to_pretty[old_cat.lower()]
            new_categories.append(new_cat)
            print(f"  '{old_cat}' -> '{new_cat}'")
        else:
            # Keep unknown categories as-is but title case them
            new_cat = old_cat.title()
            new_categories.append(new_cat)
            print(f"  '{old_cat}' -> '{new_cat}' (unknown category)")
    
    # Add any missing pretty categories
    for pretty_cat in pretty_categories:
        if pretty_cat not in new_categories:
            new_categories.append(pretty_cat)
            print(f"  Added missing category: '{pretty_cat}'")
    
    # Migrate mappings
    old_mappings = data.get('mappings', [])
    new_mappings = []
    
    print(f"\nMigrating {len(old_mappings)} mappings:")
    for mapping in old_mappings:
        old_category = mapping.get('category', '')
        if old_category.lower() in lowercase_to_pretty:
            new_category = lowercase_to_pretty[old_category.lower()]
            mapping['category'] = new_category
            print(f"  Mapping '{mapping.get('description', '')}': '{old_category}' -> '{new_category}'")
        else:
            # Keep unknown categories as-is but title case them
            new_category = old_category.title()
            mapping['category'] = new_category
            print(f"  Mapping '{mapping.get('description', '')}': '{old_category}' -> '{new_category}' (unknown)")
        
        new_mappings.append(mapping)
    
    # Update data
    data['categories'] = new_categories
    data['mappings'] = new_mappings
    
    # Add migration metadata
    if 'metadata' not in data:
        data['metadata'] = {}
    data['metadata']['migrated_to_pretty_categories'] = True
    data['metadata']['migration_backup'] = str(backup_path)
    
    # Save updated dictionary
    print(f"\nSaving migrated dictionary to {dict_path}")
    with open(dict_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print("\nMigration completed successfully!")
    print(f"Backup saved to: {backup_path}")
    print(f"Updated categories: {len(new_categories)}")
    print(f"Updated mappings: {len(new_mappings)}")

if __name__ == "__main__":
    migrate_category_dictionary() 