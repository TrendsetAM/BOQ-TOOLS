#!/usr/bin/env python3
"""
Runtime hook for BOQ-Tools executable
This hook ensures the application can find its data files when running as a standalone executable.
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path

def setup_executable_environment():
    """Setup environment for standalone executable"""
    
    # Check if running as PyInstaller executable
    if getattr(sys, 'frozen', False):
        # Get the directory where the executable is located
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller temporary directory
            bundle_dir = Path(sys._MEIPASS)
        else:
            # Fallback to executable directory
            bundle_dir = Path(sys.executable).parent
        
        # Create a user-writable directory for config files
        app_data_dir = Path.home() / 'BOQ-Tools'
        app_data_dir.mkdir(exist_ok=True)
        
        # Copy config files to user directory (always update to latest)
        config_source = bundle_dir / 'config'
        config_dest = app_data_dir / 'config'
        
        if config_source.exists():
            if not config_dest.exists():
                # First time - copy entire config directory with fresh defaults
                shutil.copytree(config_source, config_dest)
                print(f"[BOQ-Tools] First run: Copied default config files to {config_dest}")
            else:
                # Subsequent runs - only copy files that don't exist in user directory
                # This preserves user customizations while adding new default files
                for config_file in config_source.glob('*'):
                    dest_file = config_dest / config_file.name
                    if not dest_file.exists():
                        if config_file.is_file():
                            shutil.copy2(config_file, dest_file)
                            print(f"[BOQ-Tools] Added new config file: {config_file.name}")
                        elif config_file.is_dir():
                            shutil.copytree(config_file, dest_file)
                            print(f"[BOQ-Tools] Added new config directory: {config_file.name}")
                
                # Special handling: if category_dictionary.json is missing, copy the default
                category_dict_dest = config_dest / 'category_dictionary.json'
                category_dict_source = config_source / 'category_dictionary.json'
                if not category_dict_dest.exists() and category_dict_source.exists():
                    shutil.copy2(category_dict_source, category_dict_dest)
                    print(f"[BOQ-Tools] Restored missing category dictionary")
        
        # Set environment variables for the application to find config files
        os.environ['BOQ_TOOLS_CONFIG_DIR'] = str(config_dest)
        os.environ['BOQ_TOOLS_APP_DIR'] = str(app_data_dir)
        os.environ['BOQ_TOOLS_BUNDLE_DIR'] = str(bundle_dir)
        
        # Ensure the application can write log files
        log_dir = app_data_dir / 'logs'
        log_dir.mkdir(exist_ok=True)
        os.environ['BOQ_TOOLS_LOG_DIR'] = str(log_dir)

# Run the setup
setup_executable_environment() 