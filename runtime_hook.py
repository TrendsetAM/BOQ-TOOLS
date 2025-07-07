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
        
        # Copy config files to user directory if they don't exist
        config_source = bundle_dir / 'config'
        config_dest = app_data_dir / 'config'
        
        if config_source.exists() and not config_dest.exists():
            shutil.copytree(config_source, config_dest)
        
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