


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
