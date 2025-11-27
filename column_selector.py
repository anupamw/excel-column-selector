#!/usr/bin/env python3
"""
Excel Column Selector
Reads an Excel file, allows interactive column selection, and saves a filtered version.
"""

import sys
import os
import pandas as pd
import inquirer


def main():
    # Check command line arguments
    if len(sys.argv) != 2:
        print("Usage: python column_selector.py <excel_file_path>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Validate file exists
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        sys.exit(1)
    
    # Validate file extension
    if not input_file.lower().endswith(('.xlsx', '.xls')):
        print("Error: File must be an Excel file (.xlsx or .xls)")
        sys.exit(1)
    
    print(f"Reading Excel file: {input_file}")
    
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        print(f"\nFile loaded successfully! Found {len(df)} rows and {len(df.columns)} columns.\n")
        
        # Get available columns
        available_columns = list(df.columns)
        
        # Interactive column selection using checkboxes
        questions = [
            inquirer.Checkbox(
                'columns',
                message="Select columns to include (use spacebar to select, enter to confirm)",
                choices=available_columns,
                default=available_columns  # All selected by default
            )
        ]
        
        answers = inquirer.prompt(questions)
        
        # Check if user cancelled
        if answers is None or not answers['columns']:
            print("\nNo columns selected. Exiting.")
            sys.exit(0)
        
        selected_columns = answers['columns']
        
        print(f"\nSelected {len(selected_columns)} column(s): {', '.join(selected_columns)}")
        
        # Filter the dataframe
        filtered_df = df[selected_columns]
        
        # Generate output filename
        base_path = os.path.dirname(input_file)
        base_name = os.path.basename(input_file)
        name_without_ext, ext = os.path.splitext(base_name)
        output_file = os.path.join(base_path, f"{name_without_ext}_filtered{ext}")
        
        # Save the filtered Excel file
        print(f"\nSaving filtered file to: {output_file}")
        filtered_df.to_excel(output_file, index=False)
        
        print(f"✓ Success! Filtered file saved with {len(filtered_df)} rows and {len(selected_columns)} columns.")
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

