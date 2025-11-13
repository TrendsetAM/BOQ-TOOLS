#!/usr/bin/env python3
"""
Check what sheets are available in the BoQ file
"""

import pandas as pd
import os

def check_boq_sheets():
    """Check what sheets are available in the BoQ file"""
    boq_file = "examples/GRE.EEC.F.27.IT.P.18371.00.098.02 - PONTESTURA 9,69 MW_cBOQ PV_rev 9 giu (ENEMEK).xlsx"
    
    if not os.path.exists(boq_file):
        print(f"BoQ file not found: {boq_file}")
        return
    
    print(f"Loading BoQ file: {boq_file}")
    excel_data = pd.read_excel(boq_file, sheet_name=None)
    
    print(f"Available sheets: {list(excel_data.keys())}")
    
    for sheet_name, sheet_data in excel_data.items():
        print(f"\nSheet: {sheet_name}")
        print(f"  Shape: {sheet_data.shape}")
        print(f"  Columns: {list(sheet_data.columns)}")
        
        # Check if this sheet has description-like columns
        desc_columns = [col for col in sheet_data.columns if 'desc' in col.lower() or 'item' in col.lower()]
        if desc_columns:
            print(f"  Description-like columns: {desc_columns}")
        
        # Show first few rows
        print(f"  First 3 rows:")
        print(sheet_data.head(3))

if __name__ == "__main__":
    check_boq_sheets() 