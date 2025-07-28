#!/usr/bin/env python3
"""
Check the Civil Works sheet structure
"""

import pandas as pd
import os

def check_civil_works():
    """Check the Civil Works sheet structure"""
    boq_file = "examples/GRE.EEC.F.27.IT.P.18371.00.098.02 - PONTESTURA 9,69 MW_cBOQ PV_rev 9 giu (ENEMEK).xlsx"
    
    if not os.path.exists(boq_file):
        print(f"BoQ file not found: {boq_file}")
        return
    
    print(f"Loading Civil Works sheet...")
    civil_works = pd.read_excel(boq_file, sheet_name='Civil Works')
    
    print(f"Shape: {civil_works.shape}")
    print(f"Columns: {list(civil_works.columns)}")
    
    # Show first few rows
    print(f"\nFirst 5 rows:")
    print(civil_works.head())
    
    # Look for any non-null values in the first few columns
    print(f"\nChecking for non-null values in first 10 columns:")
    for i in range(min(10, len(civil_works.columns))):
        col_name = civil_works.columns[i]
        non_null_count = civil_works[col_name].notna().sum()
        print(f"  Column {i} ({col_name}): {non_null_count} non-null values")
        
        if non_null_count > 0:
            print(f"    Sample values: {civil_works[col_name].dropna().head(3).tolist()}")

if __name__ == "__main__":
    check_civil_works() 