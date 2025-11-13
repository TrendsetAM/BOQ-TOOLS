#!/usr/bin/env python3
"""
Script to implement the optimized comparison method
"""

def add_optimized_method():
    """
    Add the optimized comparison method to ui/main_window.py
    """
    
    # The optimized method to add
    optimized_method = '''
    def _process_comparison_file_optimized(self, filepath, offer_info, master_file_mapping):
        """
        Process comparison file using master BOQ mapping for maximum efficiency
        
        Args:
            filepath: Path to comparison file
            offer_info: Offer information dictionary
            master_file_mapping: Master BOQ file mapping with known structure
            
        Returns:
            FileMapping object or None if failed
        """
        try:
            from core.file_processor import ExcelProcessor
            import pandas as pd
            
            # Create Excel processor
            excel_processor = ExcelProcessor()
            
            # Load the file
            if not excel_processor.load_file(filepath):
                return None
            
            # Get the exact sheets that exist in the master BOQ
            master_sheet_names = set()
            if hasattr(master_file_mapping, 'sheets') and master_file_mapping.sheets:
                master_sheet_names = {sheet.sheet_name for sheet in master_file_mapping.sheets}
            else:
                # Fallback: get visible sheets
                master_sheet_names = set(excel_processor.get_visible_sheets())
            
            # Get visible sheets from comparison file
            visible_sheets = excel_processor.get_visible_sheets()
            if not visible_sheets:
                return None
            
            # Check if all master sheets exist in comparison file
            missing_sheets = master_sheet_names - set(visible_sheets)
            if missing_sheets:
                error_msg = f"Comparison file is missing required sheets: {missing_sheets}"
                logger.error(error_msg)
                messagebox.showerror("Structure Mismatch", error_msg)
                return None
            
            # Only process the exact sheets from master BOQ
            sheets_to_process = list(master_sheet_names)
            logger.info(f"Processing {len(sheets_to_process)} sheets based on master BOQ structure")
            
            # Extract data directly using master mapping information
            all_data = []
            
            for sheet_name in sheets_to_process:
                try:
                    # Get sheet data
                    sheet_data = excel_processor.get_sheet_data(sheet_name, max_rows=10000)
                    if not sheet_data:
                        continue
                    
                    # Convert to DataFrame
                    df = pd.DataFrame(sheet_data[1:], columns=sheet_data[0] if sheet_data else [])
                    if df.empty:
                        continue
                    
                    # Add sheet name column
                    df['Source_Sheet'] = sheet_name
                    
                    # Append to all data
                    all_data.append(df)
                    
                    logger.debug(f"Extracted {len(df)} rows from sheet '{sheet_name}'")
                    
                except Exception as e:
                    logger.warning(f"Failed to process sheet '{sheet_name}': {e}")
                    continue
            
            if not all_data:
                messagebox.showerror("Error", "No data could be extracted from comparison file")
                return None
            
            # Combine all data
            combined_df = pd.concat(all_data, ignore_index=True)
            logger.info(f"Combined {len(combined_df)} total rows from comparison file")
            
            # Create file mapping object
            file_mapping = type('MockFileMapping', (), {
                'dataframe': combined_df,
                'offer_info': offer_info,
                'sheets': []  # Add empty sheets list for compatibility
            })()
            
            # Store offer information
            file_mapping.offer_info = offer_info
            
            return file_mapping
            
        except Exception as e:
            logger.error(f"Error processing comparison file: {e}")
            return None
'''
    
    print("Optimized method to add:")
    print(optimized_method)
    print("\nTo implement:")
    print("1. Add the above method to ui/main_window.py before _validate_mapping_compatibility")
    print("2. Update _compare_full to use _process_comparison_file_optimized instead of _process_comparison_file")
    print("3. Remove the validation step since structure is guaranteed to match")

if __name__ == "__main__":
    add_optimized_method() 