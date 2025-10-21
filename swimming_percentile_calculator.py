"""
Swimming Event Percentile Calculator - Separate Reports
========================================================
This script processes multiple Excel files containing swimming event results
and creates a SEPARATE percentile report for each input file.

For swimming, lower times are better, so:
- 80th percentile = Top 20% (fastest swimmers)
- 90th percentile = Top 10% (fastest swimmers)

Requirements:
    pip install pandas openpyxl numpy

Usage:
    1. Place all your swimming event Excel files in a folder
    2. Update the INPUT_FOLDER and OUTPUT_FOLDER paths below
    3. Update RESULTS_COLUMN if your data is in a different column
    4. Run the script: python swimming_percentile_calculator.py
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')


# ===== CONFIGURATION SECTION =====
# Update these paths to match your setup
INPUT_FOLDER = "data"  # Folder containing your Excel files
OUTPUT_FOLDER = "percentile_reports"  # Folder where reports will be saved

# Column configuration
RESULTS_COLUMN = "J"  # Column letter where swimming times are stored (e.g., "A", "B", "C")
FIRST_DATA_ROW = 2  # Row number where data starts (1 = first row, 2 = second row if row 1 is header)

# Percentile configuration
# These are inverted because lower times are better in swimming
PERCENTILE_LEVELS = {
    "80th Percentile (Top 20%)": 0.20,
    "83rd Percentile (Top 17%)": 0.17,
    "85th Percentile (Top 15%)": 0.15,
    "88th Percentile (Top 12%)": 0.12,
    "90th Percentile (Top 10%)": 0.10
}
# =================================


def column_letter_to_index(column_letter):
    """Convert Excel column letter (A, B, C) to zero-based index (0, 1, 2)"""
    column_letter = column_letter.upper()
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def calculate_rank_at_percentile(data_series, percentile_value):
    """
    Calculate how many results are at or below the percentile value.
    For swimming, this represents how many swimmers achieved that time or faster.
    """
    return (data_series <= percentile_value).sum()


def process_single_sheet(sheet_data, sheet_name, results_col_index):
    """
    Process a single worksheet and calculate all percentiles.
    
    Args:
        sheet_data: DataFrame containing the sheet data
        sheet_name: Name of the worksheet
        results_col_index: Column index where results are stored
        
    Returns:
        Dictionary with percentile results or None if processing failed
    """
    try:
        # Extract the results column, skipping header rows
        results_column = sheet_data.iloc[:, results_col_index]
        
        # Skip header rows and get numeric data only
        numeric_data = pd.to_numeric(results_column.iloc[FIRST_DATA_ROW - 1:], errors='coerce')
        numeric_data = numeric_data.dropna()
        
        if len(numeric_data) == 0:
            return {
                'Event Name': sheet_name,
                'Total Results': 0,
                'Status': 'No numeric data found'
            }
        
        # Calculate all percentiles
        result_dict = {
            'Event Name': sheet_name,
            'Total Results': len(numeric_data)
        }
        
        for percentile_name, percentile_value in PERCENTILE_LEVELS.items():
            # Calculate the percentile time value
            percentile_time = np.percentile(numeric_data, percentile_value * 100)
            
            # Calculate rank (how many swimmers at or below this time)
            rank = calculate_rank_at_percentile(numeric_data, percentile_time)
            
            result_dict[percentile_name] = round(percentile_time, 2)
            result_dict[f'Rank at {percentile_name.split()[0]}'] = rank
        
        result_dict['Status'] = 'Success'
        
        return result_dict
        
    except Exception as e:
        return {
            'Event Name': sheet_name,
            'Total Results': 0,
            'Status': f'Error: {str(e)}'
        }


def process_excel_file(file_path, results_col_index, output_folder):
    """
    Process a single Excel file with multiple sheets and create a separate report.
    
    Args:
        file_path: Path to the Excel file
        results_col_index: Column index where results are stored
        output_folder: Path to output folder
        
    Returns:
        Tuple of (number of sheets processed, output file path)
    """
    results = []
    file_name = file_path.name
    file_stem = file_path.stem  # Filename without extension
    
    print(f"\nProcessing file: {file_name}")
    print("-" * 60)
    
    try:
        # Load all sheets from the Excel file
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            print(f"  - Processing sheet: {sheet_name}")
            
            # Read the sheet
            sheet_data = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            # Process the sheet
            sheet_result = process_single_sheet(sheet_data, sheet_name, results_col_index)
            
            if sheet_result:
                results.append(sheet_result)
    
    except Exception as e:
        print(f"  ERROR processing file {file_name}: {str(e)}")
        results.append({
            'Event Name': 'N/A',
            'Total Results': 0,
            'Status': f'File error: {str(e)}'
        })
        return 0, None
    
    # Create output file for this input file
    if results:
        summary_df = pd.DataFrame(results)
        
        # Reorder columns for better readability
        column_order = ['Event Name', 'Total Results']
        
        for percentile_name in PERCENTILE_LEVELS.keys():
            column_order.append(percentile_name)
            column_order.append(f'Rank at {percentile_name.split()[0]}')
        
        column_order.append('Status')
        
        summary_df = summary_df[column_order]
        
        # Create output filename
        output_file = output_folder / f"{file_stem}_percentiles.xlsx"
        
        # Write to Excel with formatting
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='Percentile Summary', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Percentile Summary']
            for idx, col in enumerate(summary_df.columns):
                max_length = max(
                    summary_df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                # Get Excel column letter
                col_letter = chr(65 + idx) if idx < 26 else chr(64 + idx // 26) + chr(65 + idx % 26)
                worksheet.column_dimensions[col_letter].width = min(max_length, 50)
        
        successful_count = sum(1 for r in results if r['Status'] == 'Success')
        print(f"  ✓ Created report: {output_file.name}")
        print(f"  ✓ Events processed: {len(results)} ({successful_count} successful)")
        
        return len(results), output_file
    
    return 0, None


def main():
    """Main function to process all Excel files and generate separate reports"""
    
    print("=" * 70)
    print("SWIMMING EVENT PERCENTILE CALCULATOR - SEPARATE REPORTS")
    print("=" * 70)
    print()
    
    # Convert column letter to index
    results_col_index = column_letter_to_index(RESULTS_COLUMN)
    
    # Get input folder path
    input_path = Path(INPUT_FOLDER)
    
    if not input_path.exists():
        print(f"ERROR: Input folder '{INPUT_FOLDER}' does not exist!")
        print("Please create the folder and add your Excel files, or update INPUT_FOLDER path.")
        return
    
    # Create output folder if it doesn't exist
    output_path = Path(OUTPUT_FOLDER)
    output_path.mkdir(exist_ok=True)
    print(f"Output folder: {OUTPUT_FOLDER}")
    print()
    
    # Find all Excel files in the input folder
    excel_files = list(input_path.glob("*.xlsx")) + list(input_path.glob("*.xls"))
    excel_files = [f for f in excel_files if not f.name.startswith('~')]  # Exclude temp files
    
    if not excel_files:
        print(f"ERROR: No Excel files found in '{INPUT_FOLDER}' folder!")
        return
    
    print(f"Found {len(excel_files)} Excel file(s) to process")
    print(f"Results column: {RESULTS_COLUMN}")
    print(f"First data row: {FIRST_DATA_ROW}")
    
    # Process all files
    total_events = 0
    total_reports = 0
    output_files = []
    
    for file_path in excel_files:
        events_processed, output_file = process_excel_file(file_path, results_col_index, output_path)
        total_events += events_processed
        if output_file:
            total_reports += 1
            output_files.append(output_file.name)
    
    # Final summary
    print()
    print("=" * 70)
    print("PROCESSING COMPLETE!")
    print("=" * 70)
    print(f"Input files processed: {len(excel_files)}")
    print(f"Output reports created: {total_reports}")
    print(f"Total events analyzed: {total_events}")
    print()
    
    if output_files:
        print("Reports created:")
        for output_file in output_files:
            print(f"  ✓ {output_file}")
        print()
        print(f"All reports saved to: {OUTPUT_FOLDER}/")


if __name__ == "__main__":
    main()