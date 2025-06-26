#!/usr/bin/env python3
"""
Create Sample Excel File
Creates a sample Excel file with Category and Description columns for testing
"""

import pandas as pd
from pathlib import Path


def create_sample_excel():
    """Create a sample Excel file with Category and Description columns"""
    
    # Sample data
    sample_data = {
        'Category': [
            'Materials',
            'Materials',
            'Materials',
            'Labor',
            'Labor',
            'Labor',
            'Equipment',
            'Equipment',
            'Services',
            'Services',
            'Overhead',
            'Overhead',
            'Civil Works',
            'Civil Works',
            'Electrical Works',
            'Electrical Works',
            'Site Costs',
            'Site Costs',
            'General Costs',
            'General Costs'
        ],
        'Description': [
            'Concrete foundation work',
            'Steel reinforcement supply',
            'Cement and aggregates',
            'Electrical installation',
            'Plumbing installation',
            'Carpentry work',
            'Crane rental for lifting',
            'Excavator for earthworks',
            'Design and engineering',
            'Testing and certification',
            'Site office setup',
            'Project management',
            'Road construction',
            'Foundation excavation',
            'MV cable installation',
            'Transformer installation',
            'Site camp construction',
            'Utilities and power supply',
            'Permits and documentation',
            'Insurance and bonds'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Save to Excel
    output_path = Path("examples/sample_categories.xlsx")
    df.to_excel(output_path, index=False)
    
    print(f"Sample Excel file created: {output_path}")
    print(f"Contains {len(df)} rows with {len(df.columns)} columns")
    print()
    print("Sample data:")
    print(df.head(10))
    print()
    print("Category distribution:")
    print(df['Category'].value_counts())
    
    return output_path


if __name__ == "__main__":
    create_sample_excel() 