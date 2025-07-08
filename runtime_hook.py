#!/usr/bin/env python3
"""
Runtime hook for BOQ-Tools executable
This hook ensures the application can find its data files when running as a standalone executable.
"""

import os
import sys
import shutil
from pathlib import Path

def setup_user_config():
    """Setup user config directory and copy missing files from bundled resources."""
    if not getattr(sys, 'frozen', False):
        return  # Only run when frozen (executable)
    
    # Get user config directory
    user_config_dir = Path.home() / "AppData" / "Roaming" / "BOQ-TOOLS"
    user_config_dir.mkdir(parents=True, exist_ok=True)
    
    # Get bundled resources directory
    bundle_dir = Path(getattr(sys, '_MEIPASS', ''))
    bundled_config_dir = bundle_dir / "config"
    
    # Files to copy if missing (excluding unused files)
    config_files = [
        "category_dictionary.json",
        "canonical_mappings.json", 
        "boq_settings.json"
    ]
    
    for filename in config_files:
        user_file = user_config_dir / filename
        bundled_file = bundled_config_dir / filename
        
        # Only copy if file doesn't exist in user directory
        if not user_file.exists() and bundled_file.exists():
            try:
                shutil.copy2(bundled_file, user_file)
                print(f"Copied {filename} to user config directory")
            except Exception as e:
                print(f"Error copying {filename}: {e}")
                sys.exit(1)

# Run setup when module is imported
setup_user_config() 