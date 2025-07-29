#!/usr/bin/env python3
"""
BOQ Tools - Main Application
Complete application integration with GUI and CLI modes
"""

import sys
import os
import argparse
import logging
import signal
import threading
import time
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
import traceback
import json

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Core components
from core.file_processor import ExcelProcessor
from core.sheet_classifier import SheetClassifier
from core.column_mapper import ColumnMapper
from core.row_classifier import RowClassifier
from core.validator import DataValidator
from core.mapping_generator import MappingGenerator, FileMapping

# UI components
try:
    from ui.main_window import MainWindow
    from ui.preview_dialog import show_preview_dialog
    from ui.progress_dialog import show_progress_dialog, ProgressDialog
    from ui.settings_dialog import show_settings_dialog
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

# Utils
from utils.config import get_config, BOQConfig, ensure_default_config, get_user_config_path
from utils.export import ExcelExporter
from utils.logger import setup_logging

# Application settings
APP_NAME = "BOQ Tools"
APP_VERSION = "1.0.0"
DEFAULT_CONFIG_FILE = "config/boq_settings.json"
DEFAULT_LOG_FILE = "logs/boq_tools.log"


class BOQApplicationController:
    """
    Main application controller coordinating all components
    """
    
    def __init__(self, config_file: Optional[Path] = None, log_file: Optional[Path] = None):
        """
        Initialize the application controller
        
        Args:
            config_file: Path to configuration file
            log_file: Path to log file
        """
        self.config_file = config_file or Path(DEFAULT_CONFIG_FILE)
        self.log_file = log_file or Path(DEFAULT_LOG_FILE)
        
        # Initialize components
        self.config: Optional[BOQConfig] = None
        self.logger: Optional[logging.Logger] = None
        self.processor: Optional[ExcelProcessor] = None
        self.sheet_classifier: Optional[SheetClassifier] = None
        self.column_mapper: Optional[ColumnMapper] = None
        self.row_classifier: Optional[RowClassifier] = None
        self.validator: Optional[DataValidator] = None
        self.mapping_generator: Optional[MappingGenerator] = None
        self.exporter: Optional[ExcelExporter] = None
        
        # Application state
        self.is_running = False
        self.current_files = {}
        self.settings = {}
        self.auto_save_timer = None
        
        # Comparison workflow state
        self.comparison_processor = None
        self.master_file_mapping = None
        self.is_comparison_mode = False
        
        # Threading
        self.processing_lock = threading.Lock()
        self.shutdown_event = threading.Event()
        
        # Initialize application
        self._initialize_application()
    
    def _initialize_application(self):
        """Initialize all application components"""
        try:
            # Setup logging
            self._setup_logging()
            
            # Load configuration
            self._load_configuration()
            
            # Initialize core components
            self._initialize_core_components()
            
            # Load settings
            self._load_settings()
            
            # Setup signal handlers
            self._setup_signal_handlers()
            
            assert self.logger is not None
            self.logger.info(f"{APP_NAME} v{APP_VERSION} initialized successfully")
            
        except Exception as e:
            print(f"Failed to initialize application: {e}")
            traceback.print_exc()
            sys.exit(1)
    
    def _setup_logging(self):
        """Setup logging system"""
        # Create log directory
        self.log_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Setup logging
        self.logger = setup_logging(
            log_file=self.log_file,
            level=logging.DEBUG,  # Temporarily set to DEBUG to see offer info logs
            console_output=True
        )
        
        self.logger.info("Logging system initialized")
    
    def _load_configuration(self):
        """Load application configuration"""
        assert self.logger is not None
        try:
            self.config = get_config()
            self.logger.info("Configuration loaded successfully")
        except Exception as e:
            self.logger.error(f"Failed to load configuration: {e}")
            raise
    
    def _initialize_core_components(self):
        """Initialize all core processing components"""
        assert self.logger is not None
        try:
            # Get max_header_rows from user preferences
            max_header_rows = self.settings.get("user_preferences", {}).get("processing_thresholds", {}).get("max_header_rows", 20)
            
            # Initialize processors
            self.processor = ExcelProcessor()
            self.sheet_classifier = SheetClassifier()
            self.column_mapper = ColumnMapper(max_header_rows=max_header_rows)
            self.row_classifier = RowClassifier()
            self.validator = DataValidator()
            self.mapping_generator = MappingGenerator()
            self.exporter = ExcelExporter()
            
            self.logger.info("Core components initialized successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to initialize core components: {e}")
            raise
    
    def _load_settings(self):
        """Load application settings"""
        assert self.logger is not None
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                self.logger.info("Settings loaded successfully")
            else:
                self.settings = {}
                self.logger.info("No settings file found, using defaults")
        except Exception as e:
            self.logger.error(f"Failed to load settings: {e}")
            self.settings = {}
    
    def _save_settings(self):
        """Save application settings"""
        assert self.logger is not None
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2, ensure_ascii=False)
            self.logger.debug("Settings saved successfully")
        except Exception as e:
            self.logger.error(f"Failed to save settings: {e}")
    
    def _setup_signal_handlers(self):
        """Setup signal handlers for graceful shutdown"""
        def signal_handler(signum, frame):
            assert self.logger is not None
            self.logger.info(f"Received signal {signum}, initiating shutdown")
            self.shutdown()
        
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
    
    def process_file(self, file_path: Path, progress_callback: Optional[Callable] = None, sheet_filter: Optional[List[str]] = None, sheet_types: Optional[Dict[str, str]] = None) -> FileMapping:
        """
        Process a single Excel file through the complete pipeline
        
        Args:
            file_path: Path to Excel file
            progress_callback: Optional progress callback function
            sheet_filter: Optional list of sheet names to process (only these will be processed)
            sheet_types: Optional dict mapping sheet names to their user-selected type (e.g., 'BOQ', 'Info', 'Ignore')
        
        Returns:
            Complete processing results
        """
        assert self.logger is not None
        assert self.processor is not None
        assert self.sheet_classifier is not None
        assert self.column_mapper is not None
        assert self.row_classifier is not None
        assert self.validator is not None
        assert self.mapping_generator is not None

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        # Use an absolute path for the dictionary key to ensure consistency
        abs_filepath_str = str(file_path.resolve())
        
        self.logger.info(f"Processing file: {file_path}")
        
        try:
            # Load the file first, outside the context manager
            self.processor.load_file(file_path)

            # Step 1: Use the processor as a context manager to ensure cleanup
            with self.processor as processor_instance:
                if progress_callback:
                    progress_callback(10, "File loaded, analyzing...")

                file_info = processor_instance.get_file_info()
                sheet_data = processor_instance.get_all_sheets_data()
                
                if not sheet_data:
                    raise ValueError("No data found in any sheets.")

            # Filter sheets if a filter is provided
            if sheet_filter is not None:
                sheet_data = {name: data for name, data in sheet_data.items() if name in sheet_filter}

            if progress_callback:
                progress_callback(20, "Classifying sheets...")
            
            # The rest of the processing can happen outside the 'with' block
            # as the data has been loaded into memory.
            
            sheet_classifications = {
                name: self.sheet_classifier.classify_sheet(data, name) 
                for name, data in sheet_data.items()
            }
            if progress_callback: progress_callback(30, "Sheets classified")

            column_mapping_results = {
                name: self.column_mapper.process_sheet_mapping(data)
                for name, data in sheet_data.items()
            }
            if progress_callback: progress_callback(50, "Columns mapped")

            column_mappings_dict = {
                name: {m.column_index: m.mapped_type for m in result.mappings}
                for name, result in column_mapping_results.items()
            }
            
            row_classifications = {
                name: self.row_classifier.classify_rows(data, column_mappings_dict.get(name, {}), name)
                for name, data in sheet_data.items()
            }
            if progress_callback: progress_callback(70, "Rows classified")

            row_classifications_dict = {
                name: {rc.row_index: rc.row_type.value for rc in result.classifications}
                for name, result in row_classifications.items()
            }

            validation_results = {
                name: self.validator.validate_sheet(
                    data, column_mappings_dict.get(name, {}), row_classifications_dict.get(name, {})
                )
                for name, data in sheet_data.items()
            }
            if progress_callback: progress_callback(90, "Data validated")

            processor_results = {
                'file_info': file_info,
                'sheet_data': sheet_data,
                'sheet_classifications': sheet_classifications,
                'column_mappings': column_mapping_results,
                'row_classifications': row_classifications,
                'validation_results': validation_results,
                'sheet_types': sheet_types or {name: 'BOQ' for name in sheet_data.keys()}
            }
            
            file_mapping = self.mapping_generator.generate_file_mapping(processor_results)
            if progress_callback: progress_callback(100, "Processing complete")

            # Add column mapper reference to file mapping for UI learning functionality
            file_mapping.column_mapper = self.column_mapper

            self.current_files[abs_filepath_str] = {
                'file_mapping': file_mapping,
                'processor_results': processor_results,
                'processing_time': time.time()
            }
            
            self.logger.info(f"File processing completed: {file_path}")
            return file_mapping
                
        except Exception as e:
            self.logger.error(f"Error processing file {file_path}: {e}", exc_info=True)
            raise
    
    def export_file(self, file_path: str, export_path: Path, format_type: str) -> bool:
        """
        Export processed file using the appropriate format.
        
        Args:
            file_path: Key of the processed file (should be absolute path).
            export_path: Export destination path.
            format_type: Export format type ('normalized_excel', 'summary_excel', 'json', 'csv').
            
        Returns:
            True if export is successful.
        """
        # The key in current_files is an absolute path string.
        # Ensure the incoming file_path matches this.
        processed_file_key = str(Path(file_path).resolve())

        if processed_file_key not in self.current_files:
            # For debugging, let's see what keys are available
            available_keys = list(self.current_files.keys())
            assert self.logger is not None
            self.logger.error(f"Export failed: '{processed_file_key}' not found. Available files: {available_keys}")
            raise ValueError(f"File not found in processed files: {file_path}")
        
        try:
            assert self.logger is not None
            assert self.exporter is not None
            file_data = self.current_files[processed_file_key]
            file_mapping = file_data['file_mapping']
            
            # This now correctly dispatches to the exporter
            return self.exporter.export_data(file_mapping, export_path, format_type)
                
        except Exception as e:
            assert self.logger is not None
            self.logger.error(f"Error exporting file {file_path}: {e}", exc_info=True)
            return False
    
    def batch_process(self, file_paths: List[Path], output_dir: Path) -> Dict[str, bool]:
        """
        Process multiple files in batch
        
        Args:
            file_paths: List of file paths to process
            output_dir: Output directory for exports
            
        Returns:
            Dictionary mapping file paths to success status
        """
        results = {}
        
        for file_path in file_paths:
            try:
                assert self.logger is not None
                self.logger.info(f"Processing file: {file_path}")
                
                # Process file
                file_mapping = self.process_file(file_path)
                
                # Export file
                export_path = output_dir / f"{file_path.stem}_processed.xlsx"
                success = self.export_file(str(file_path), export_path, 'normalized_excel')
                
                results[str(file_path)] = success
                
            except Exception as e:
                assert self.logger is not None
                self.logger.error(f"Failed to process {file_path}: {e}")
                results[str(file_path)] = False
        
        return results
    
    def get_processing_status(self) -> Dict[str, Any]:
        """Get current processing status"""
        return {
            'files_processed': len(self.current_files),
            'is_processing': self.processing_lock.locked(),
            'settings_loaded': bool(self.settings),
            'components_initialized': all([
                self.processor, self.sheet_classifier, self.column_mapper,
                self.row_classifier, self.validator, self.mapping_generator
            ])
        }
    
    def get_current_dataframe(self):
        """Get the current DataFrame from the most recently processed file"""
        try:
            if not self.current_files:
                return None
            
            # Get the most recently processed file
            latest_file_key = max(self.current_files.keys(), 
                                key=lambda k: self.current_files[k].get('processing_time', 0))
            
            file_data = self.current_files[latest_file_key]
            file_mapping = file_data.get('file_mapping')
            
            if not file_mapping:
                return None
            
            # Try to get DataFrame from file mapping
            if hasattr(file_mapping, 'dataframe'):
                return file_mapping.dataframe
            elif hasattr(file_mapping, 'get_processed_dataframe'):
                return file_mapping.get_processed_dataframe()
            else:
                # Create DataFrame from the file mapping data
                return self._create_dataframe_from_mapping(file_mapping)
                
        except Exception as e:
            assert self.logger is not None
            self.logger.error(f"Error getting current DataFrame: {e}")
            return None
    
    def _create_dataframe_from_mapping(self, file_mapping):
        """Create a DataFrame from file mapping data"""
        try:
            import pandas as pd
            import logging
            logger = logging.getLogger(__name__)
            
            rows = []
            for sheet in getattr(file_mapping, 'sheets', []):
                if getattr(sheet, 'sheet_type', 'BOQ') != 'BOQ':
                    continue
                
                col_headers = [cm.mapped_type for cm in getattr(sheet, 'column_mappings', [])]
                sheet_name = sheet.sheet_name
                
                # DEBUG: Log column mappings for this sheet
                logger.info(f"Sheet '{sheet_name}' column mappings: {col_headers}")
                logger.info(f"Sheet '{sheet_name}' ignore columns: {[cm.mapped_type for cm in getattr(sheet, 'column_mappings', []) if cm.mapped_type == 'ignore']}")
                
                # Get row validity for this sheet
                sheet_validity = getattr(file_mapping, 'row_validity', {}).get(sheet_name, {})
                
                for rc in getattr(sheet, 'row_classifications', []):
                    # Only include valid rows
                    if not sheet_validity.get(rc.row_index, True):
                        continue
                    
                    row_data = getattr(rc, 'row_data', None)
                    if row_data is None and hasattr(sheet, 'sheet_data'):
                        try:
                            row_data = sheet.sheet_data[rc.row_index]
                        except Exception:
                            row_data = None
                    
                    if row_data is None:
                        row_data = []
                    
                    row_dict = {}
                    # Add Source_Sheet column for each row
                    row_dict['Source_Sheet'] = sheet_name
                    
                    for cm in sheet.column_mappings:
                        mapped_type = getattr(cm, 'mapped_type', None)
                        if not mapped_type:
                            continue
                        # Skip columns mapped to 'ignore' - they shouldn't be in the DataFrame
                        if mapped_type == 'ignore':
                            continue
                        idx = cm.column_index
                        row_dict[mapped_type] = row_data[idx] if idx < len(row_data) else ''
                    

                    
                    # Only add essential columns that are not 'ignore'
                    essential_columns = ['description', 'quantity', 'unit_price', 'total_price', 'unit', 'code', 'scope', 'manhours', 'wage', 'category']
                    for mt in col_headers:
                        if mt not in row_dict and mt in essential_columns:
                            row_dict[mt] = ''
                    
                    rows.append(row_dict)
            
            if rows:
                df = pd.DataFrame(rows)
                
                # DEBUG: Log DataFrame columns before normalization
                logger.info(f"DataFrame columns before normalization: {list(df.columns)}")
                
                # Normalize column names to be case-insensitive
                # Map common variations to standard names
                column_mapping = {
                    'description': 'Description',
                    'Description': 'Description',
                    'DESCRIPTION': 'Description',
                    'category': 'Category',
                    'Category': 'Category',
                    'CATEGORY': 'Category',
                    'code': 'code',
                    'unit': 'unit',
                    'quantity': 'quantity',
                    'unit_price': 'unit_price',
                    'total_price': 'total_price'
                }
                
                # Rename columns to standard names
                df_renamed = df.copy()
                for col in df.columns:
                    if col.lower() in column_mapping:
                        new_name = column_mapping[col.lower()]
                        if col != new_name:
                            df_renamed = df_renamed.rename(columns={col: new_name})
                
                # DEBUG: Log DataFrame columns after normalization
                logger.info(f"DataFrame columns after normalization: {list(df_renamed.columns)}")
                
                # Reorder columns to the correct sequence (Source_Sheet is now added during DataFrame creation)
                desired_order = ['Source_Sheet', 'code', 'Category', 'Description', 'unit', 'quantity', 'unit_price', 'total_price', 'manhours', 'wage']
                
                # Add any missing columns with empty values
                for col in desired_order:
                    if col not in df_renamed.columns:
                        df_renamed[col] = ''
                
                # Reorder columns to match desired sequence
                available_columns = [col for col in desired_order if col in df_renamed.columns]
                remaining_columns = [col for col in df_renamed.columns if col not in desired_order]
                final_columns = available_columns + remaining_columns
                
                df_renamed = df_renamed[final_columns]
                
                # DEBUG: Log final column order
                logger.info(f"Final DataFrame columns in order: {list(df_renamed.columns)}")
                
                return df_renamed
            else:
                return None
                
        except Exception as e:
            assert self.logger is not None
            self.logger.error(f"Error creating DataFrame from mapping: {e}")
            return None
    
    def update_settings(self, new_settings: Dict[str, Any]):
        """Update application settings"""
        self.settings.update(new_settings)
        self._save_settings()
        
        # Update ColumnMapper with new max_header_rows setting if it changed
        if self.column_mapper:
            max_header_rows = self.settings.get("user_preferences", {}).get("processing_thresholds", {}).get("max_header_rows", 20)
            self.column_mapper.max_header_rows = max_header_rows
        
        assert self.logger is not None
        self.logger.info("Settings updated and saved")
    
    def auto_save(self):
        """Perform auto-save operation"""
        try:
            self._save_settings()
            assert self.logger is not None
            self.logger.debug("Auto-save completed")
        except Exception as e:
            assert self.logger is not None
            self.logger.error(f"Auto-save failed: {e}")
    
    def start_auto_save(self, interval_seconds: int = 300):
        """Start auto-save timer"""
        def auto_save_worker():
            while not self.shutdown_event.is_set():
                time.sleep(interval_seconds)
                if not self.shutdown_event.is_set():
                    self.auto_save()
        
        self.auto_save_timer = threading.Thread(target=auto_save_worker, daemon=True)
        self.auto_save_timer.start()
        assert self.logger is not None
        self.logger.info(f"Auto-save started with {interval_seconds}s interval")
    
    def stop_auto_save(self):
        """Stop auto-save timer"""
        if self.auto_save_timer:
            self.shutdown_event.set()
            self.auto_save_timer.join(timeout=5)
            assert self.logger is not None
            self.logger.info("Auto-save stopped")
    
    def shutdown(self):
        """Graceful application shutdown"""
        if not self.is_running:
            return
        
        assert self.logger is not None
        self.logger.info("Initiating application shutdown")
        self.is_running = False
        
        # Stop auto-save
        self.stop_auto_save()
        
        # Save final state
        self._save_settings()
        
        # Clear current files
        self.current_files.clear()
        
        assert self.logger is not None
        self.logger.info("Application shutdown completed")
    
    def start_comparison_workflow(self, master_file_mapping: FileMapping):
        """
        Start comparison workflow with master file mapping
        
        Args:
            master_file_mapping: FileMapping for the master BoQ
        """
        assert self.logger is not None
        self.logger.info("Starting comparison workflow")
        
        self.is_comparison_mode = True
        self.master_file_mapping = master_file_mapping
        
        # Initialize comparison processor
        from core.comparison_engine import ComparisonProcessor
        self.comparison_processor = ComparisonProcessor()
        
        self.logger.info("Comparison workflow initialized")
    
    def process_comparison_file(self, file_path: Path, offer_info: Dict[str, Any]) -> Optional[FileMapping]:
        """
        Process comparison file in comparison workflow
        
        Args:
            file_path: Path to comparison file
            offer_info: Offer information for comparison
            
        Returns:
            FileMapping for comparison file or None if failed
        """
        assert self.logger is not None
        
        if not self.is_comparison_mode:
            self.logger.error("Not in comparison mode")
            return None
        
        try:
            # Process the comparison file using normal file processing
            comparison_mapping = self.process_file(file_path)
            
            # Store offer information
            comparison_mapping.offer_info = offer_info
            
            self.logger.info(f"Comparison file processed: {file_path}")
            return comparison_mapping
            
        except Exception as e:
            self.logger.error(f"Error processing comparison file {file_path}: {e}")
            return None
    
    def get_comparison_processor(self) -> Optional[Any]:
        """Get the current comparison processor"""
        return self.comparison_processor
    
    def end_comparison_workflow(self):
        """End comparison workflow and reset state"""
        assert self.logger is not None
        self.logger.info("Ending comparison workflow")
        
        self.is_comparison_mode = False
        self.master_file_mapping = None
        self.comparison_processor = None


class BOQApplication:
    """
    Main application class handling GUI and CLI modes
    """
    
    def __init__(self, config_file: Optional[Path] = None, log_file: Optional[Path] = None):
        """Initialize the application"""
        self.controller = BOQApplicationController(config_file, log_file)
        self.gui_mode = False
        self.main_window = None
    
    def run_gui(self):
        """Run the application in GUI mode"""
        if not GUI_AVAILABLE:
            print("GUI components not available. Running in CLI mode.")
            return self.run_cli()
        
        try:
            self.gui_mode = True
            self.main_window = MainWindow(self.controller)
            self._connect_ui_to_controller()
            self.main_window.run()
            
        except Exception as e:
            assert self.controller.logger is not None
            self.controller.logger.error(f"Failed to run GUI: {e}", exc_info=True)
            traceback.print_exc()
            # Fallback to CLI
            print("GUI failed to start, falling back to CLI mode.")
            self.run_cli()
    
    def run_cli(self, args: Optional[argparse.Namespace] = None):
        """Run the application in CLI mode"""
        try:
            self.controller.is_running = True
            
            if args is None:
                # Interactive CLI mode
                self._run_interactive_cli()
            else:
                # Command-line argument mode
                self._run_argument_cli(args)
                
        except Exception as e:
            assert self.controller.logger is not None
            self.controller.logger.error(f"CLI error: {e}")
            traceback.print_exc()
        finally:
            self.controller.shutdown()
    
    def _run_interactive_cli(self):
        """Run interactive CLI mode"""
        print(f"Welcome to {APP_NAME} v{APP_VERSION}")
        print("Interactive CLI mode. Type 'help' for commands.")
        
        while self.controller.is_running:
            try:
                command = input("BOQ> ").strip().lower()
                
                if command == 'quit' or command == 'exit':
                    break
                elif command == 'help':
                    self._show_cli_help()
                elif command == 'status':
                    self._show_status()
                elif command.startswith('process '):
                    file_path = command[8:].strip()
                    self._process_file_cli(Path(file_path))
                elif command.startswith('export '):
                    parts = command[7:].split()
                    if len(parts) >= 2:
                        file_key = parts[0]
                        export_path = Path(parts[1])
                        format_type = parts[2] if len(parts) > 2 else 'normalized_excel'
                        self._export_file_cli(file_key, export_path, format_type)
                elif command == 'list':
                    self._list_processed_files()
                elif command == 'clear':
                    self.controller.current_files.clear()
                    print("Cleared processed files")
                else:
                    print("Unknown command. Type 'help' for available commands.")
                    
            except KeyboardInterrupt:
                print("\nUse 'quit' to exit")
            except Exception as e:
                print(f"Error: {e}")
    
    def _run_argument_cli(self, args: argparse.Namespace):
        """Run CLI mode with command-line arguments"""
        if args.file:
            # Process single file
            file_path = Path(args.file)
            if not file_path.exists():
                print(f"Error: File not found: {file_path}")
                return
            
            try:
                print(f"Processing file: {file_path}")
                file_mapping = self.controller.process_file(file_path)
                print("Processing completed successfully")
                
                if args.export:
                    export_path = Path(args.export)
                    success = self.controller.export_file(
                        str(file_path), 
                        export_path, 
                        args.format or 'normalized_excel'
                    )
                    if success:
                        print(f"File exported to: {export_path}")
                    else:
                        print("Export failed")
                        
            except Exception as e:
                print(f"Error processing file: {e}")
                return
        
        elif args.batch:
            # Batch processing
            input_dir = Path(args.batch)
            output_dir = Path(args.output) if args.output else input_dir / "processed"
            
            if not input_dir.exists():
                print(f"Error: Input directory not found: {input_dir}")
                return
            
            # Find Excel files
            excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls"))
            
            if not excel_files:
                print(f"No Excel files found in: {input_dir}")
                return
            
            print(f"Found {len(excel_files)} Excel files")
            results = self.controller.batch_process(excel_files, output_dir)
            
            # Show results
            successful = sum(1 for success in results.values() if success)
            print(f"Processing completed: {successful}/{len(excel_files)} files successful")
            
            for file_path, success in results.items():
                status = "✓" if success else "✗"
                print(f"  {status} {Path(file_path).name}")
    
    def _show_cli_help(self):
        """Show CLI help"""
        print("Available commands:")
        print("  process <file>     - Process a single Excel file")
        print("  export <key> <path> [format] - Export processed file")
        print("  list               - List processed files")
        print("  status             - Show processing status")
        print("  clear              - Clear processed files")
        print("  help               - Show this help")
        print("  quit/exit          - Exit application")
        print("\nExport formats: normalized_excel, summary_excel, json, csv")
    
    def _show_status(self):
        """Show application status"""
        status = self.controller.get_processing_status()
        print(f"Files processed: {status['files_processed']}")
        print(f"Currently processing: {status['is_processing']}")
        print(f"Settings loaded: {status['settings_loaded']}")
        print(f"Components ready: {status['components_initialized']}")
    
    def _process_file_cli(self, file_path: Path):
        """Process file in CLI mode"""
        try:
            print(f"Processing {file_path}...")
            
            # Create progress callback for CLI
            def progress_callback(percentage, message):
                print(f"  {percentage}% - {message}")
            
            file_mapping = self.controller.process_file(file_path, progress_callback)
            print("Processing completed successfully")
            
            # Show summary
            summary = file_mapping.processing_summary
            print(f"  Sheets: {summary.successful_sheets} successful, {summary.partial_sheets} partial")
            print(f"  Confidence: {summary.average_confidence:.1%}")
            print(f"  Errors: {summary.total_validation_errors}")
            
        except Exception as e:
            print(f"Error: {e}")
    
    def _export_file_cli(self, file_key: str, export_path: Path, format_type: str):
        """Export file in CLI mode"""
        try:
            success = self.controller.export_file(file_key, export_path, format_type)
            if success:
                print(f"File exported successfully to: {export_path}")
            else:
                print("Export failed")
        except Exception as e:
            print(f"Error: {e}")
    
    def _list_processed_files(self):
        """List processed files"""
        if not self.controller.current_files:
            print("No files processed")
            return
        
        print("Processed files:")
        for file_key, file_data in self.controller.current_files.items():
            file_path = Path(file_key)
            processing_time = file_data.get('processing_time', 0)
            print(f"  {file_path.name} (processed at {time.ctime(processing_time)})")
    
    def _connect_ui_to_controller(self):
        """Connect UI components to controller"""
        # This would be implemented to connect the UI to the controller
        # For now, it's a placeholder
        pass


def create_parser() -> argparse.ArgumentParser:
    """Create command-line argument parser"""
    parser = argparse.ArgumentParser(
        description=f"{APP_NAME} - Excel BOQ Processing Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --gui                    # Run GUI mode
  %(prog)s --file data.xlsx         # Process single file
  %(prog)s --file data.xlsx --export output.xlsx  # Process and export
  %(prog)s --batch ./input --output ./processed   # Batch process
  %(prog)s                          # Interactive CLI mode
        """
    )
    
    # Mode selection
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument('--gui', action='store_true', help='Run in GUI mode')
    mode_group.add_argument('--file', type=str, help='Process single Excel file')
    mode_group.add_argument('--batch', type=str, help='Batch process directory of Excel files')
    
    # Output options
    parser.add_argument('--output', type=str, help='Output directory for batch processing')
    parser.add_argument('--export', type=str, help='Export path for single file processing')
    parser.add_argument('--format', choices=['normalized_excel', 'summary_excel', 'json', 'csv'], 
                       default='normalized_excel', help='Export format')
    
    # Configuration
    parser.add_argument('--config', type=str, help='Configuration file path')
    parser.add_argument('--log', type=str, help='Log file path')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose logging')
    
    return parser


def main():
    """Main application entry point"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup paths
    config_file = Path(args.config) if args.config else None
    log_file = Path(args.log) if args.log else None
    
    # Create application
    app = BOQApplication(config_file, log_file)
    
    try:
        # Determine mode and run
        # Check if running as executable (PyInstaller sets sys.frozen)
        is_executable = getattr(sys, 'frozen', False)
        
        if args.gui:
            app.run_gui()
        elif args.file or args.batch:
            app.run_cli(args)
        elif is_executable:
            # When running as executable, default to GUI mode
            app.run_gui()
        else:
            # Interactive CLI mode (only when running as Python script)
            app.run_cli()
            
    except KeyboardInterrupt:
        print("\nApplication interrupted by user")
    except Exception as e:
        print(f"Application error: {e}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        # Ensure cleanup
        if hasattr(app, 'controller'):
            app.controller.shutdown()


if __name__ == "__main__":
    # At startup, ensure user-writable config files exist
    boq_settings_path = ensure_default_config(
        'boq_settings.json',
        os.path.join(os.path.dirname(__file__), 'config', 'boq_settings.json'),
        default_data={}
    )
    canonical_mappings_path = ensure_default_config(
        'canonical_mappings.json',
        os.path.join(os.path.dirname(__file__), 'config', 'canonical_mappings.json'),
        default_data={}
    )

    # Example: Load settings
    with open(boq_settings_path, 'r', encoding='utf-8') as f:
        settings = json.load(f)

    main() 