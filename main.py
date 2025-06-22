#!/usr/bin/env python3
"""
BOQ Tools - Bill of Quantities Excel Processor
Main application entry point
"""

import sys
import logging
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from core.boq_processor import BOQProcessor
from ui.main_window import MainWindow
from utils.logger import setup_logging
from utils.config import validate_and_log_config


def main():
    """Main application entry point"""
    # Setup logging
    setup_logging()
    logger = logging.getLogger(__name__)
    
    try:
        logger.info("Starting BOQ Tools application")
        
        # Validate configuration
        if not validate_and_log_config():
            logger.error("Configuration validation failed. Exiting.")
            sys.exit(1)
        
        # Initialize the BOQ processor
        processor = BOQProcessor()
        
        # Launch the main UI
        app = MainWindow(processor)
        app.run()
        
    except Exception as e:
        logger.error(f"Application failed to start: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 