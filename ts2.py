import pandas as pd
import os
import re

def time_to_seconds(time_str):
    if pd.isna(time_str):
        return None
    
    time_str = str(time_str).strip()
    
    if not any(char.isdigit() for char in time_str):
        return None
    
    try:
        if ':' in time_str:
            parts = time_str.split(':')
            minutes = float(parts[0])
            seconds = float(parts[1])
            return minutes * 60 + seconds
        else:
            return float(time_str)
    except (ValueError, IndexError):
        return None

def seconds_to_time(seconds):
    if seconds is None:
        return None
    
    minutes = int(seconds // 60)
    secs = seconds % 60
    
    if minutes > 0:
        return f"{minutes}:{secs:05.2f}"
    else:
        return f"{secs:.2f}"

def parse_filename(filename):
    # Pattern: CAN_2025_COURSE_SEX_XX-YY
    pattern = r'CAN_2025_(SCM|LCM)_(Men|Women)_(\d+)-(\d+)'
    match = re.search(pattern, filename)
    
    if match:
        course = match.group(1)
        sex = match.group(2)
        age_start = match.group(3)
        age_end = match.group(4)
        return course, sex, age_start, age_end
    return None, None, None, None

def find_closest_rank(df, target_time_seconds, rank_col, time_col):
    """Find the rank with the time closest to the target time"""
    if target_time_seconds is None:
        return None
    
    # Convert all times to seconds
    df_copy = df.copy()
    df_copy['time_seconds'] = df_copy[time_col].apply(time_to_seconds)
    
    # Remove rows with invalid times
    df_copy = df_copy[df_copy['time_seconds'].notna()]
    
    if df_copy.empty:
        return None
    
    # Calculate absolute difference from target
    df_copy['diff'] = abs(df_copy['time_seconds'] - target_time_seconds)
    
    # Find the row with minimum difference
    closest_row = df_copy.loc[df_copy['diff'].idxmin()]
    
    return closest_row[rank_col]

def process_file(filepath):
    """Process a single Excel file and extract 50th place times"""
    print(f"\nProcessing: {os.path.basename(filepath)}")
    
    xls = pd.ExcelFile(filepath)
    results = []
    
    # Process each sheet (event) - maintain order
    for sheet_name in xls.sheet_names:
        # Skip any event with "Lap" in the name
        if "Lap" in sheet_name:
            continue
            
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Find the row where rank (column M, index 12) equals 50
        if len(df.columns) > 12 and len(df.columns) > 9:
            rank_col = df.columns[12]
            time_col = df.columns[9]
            
            # Find row where rank = 50
            rank_50_rows = df[df[rank_col] == 50]
            
            if not rank_50_rows.empty:
                time_value = rank_50_rows.iloc[0][time_col]
                
                # Convert time to seconds for calculation
                time_seconds = time_to_seconds(time_value)
                
                if time_seconds is not None:
                    # Calculate times for different percentages
                    percentages = [10, 11, 11.5, 12, 12.5]
                    adjusted_times = {}
                    
                    for pct in percentages:
                        multiplier = 1 + (pct / 100)
                        adjusted_seconds = time_seconds * multiplier
                        adjusted_time_str = seconds_to_time(adjusted_seconds)
                        closest_rank = find_closest_rank(df, adjusted_seconds, rank_col, time_col)
                        
                        adjusted_times[pct] = {
                            'time': adjusted_time_str,
                            'rank': closest_rank
                        }
                    
                    results.append({
                        'event': sheet_name,
                        '50th_time': str(time_value),
                        'adjusted_times': adjusted_times
                    })
    
    return results

def create_simplified_output(men_data, women_data, men_age_groups, women_age_groups, 
                             men_event_order, women_event_order):
    """Create a simplified output with one tab per age group showing just the percentage times"""
    
    simplified_file = 'swim_rankings_simplified.xlsx'
    
    try:
        with pd.ExcelWriter(simplified_file, engine='openpyxl') as writer:
            # Process each men's age group
            for age_group_label in men_age_groups:
                rows = []
            
                for event_name in men_event_order:
                    if age_group_label in men_data[event_name]:
                        row = {'Event': event_name}
                    
                        # Add percentage times (without ranks)
                        for pct in [10, 11, 11.5, 12, 12.5]:
                            pct_label = str(pct).replace('.', '_')
                            row[f'+{pct}%'] = men_data[event_name][age_group_label]['adjusted_times'][pct]['time']
                    
                        rows.append(row)
            
                if rows:
                    df = pd.DataFrame(rows)
                    # Clean up sheet name (Excel has 31 char limit and special char restrictions)
                    sheet_name = age_group_label[:31]
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
        
            # Process each women's age group
            for age_group_label in women_age_groups:
                rows = []
            
                for event_name in women_event_order:
                    if age_group_label in women_data[event_name]:
                        row = {'Event': event_name}
                    
                        # Add percentage times (without ranks)
                        for pct in [10, 11, 11.5, 12, 12.5]:
                            pct_label = str(pct).replace('.', '_')
                            row[f'+{pct}%'] = women_data[event_name][age_group_label]['adjusted_times'][pct]['time']
                    
                        rows.append(row)
            
                if rows:
                    df = pd.DataFrame(rows)
                    # Clean up sheet name (Excel has 31 char limit and special char restrictions)
                    sheet_name = age_group_label[:31]
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        print(f"Simplified results exported to: {simplified_file}")
    
    except PermissionError:
        print(f"\nERROR: Cannot write to '{simplified_file}'")
        print(f"Please close the file if it's open in Excel and try again.")

def main():
    data_folder = 'data'
    
    if not os.path.exists(data_folder):
        print(f"Error: '{data_folder}' folder not found!")
        return
    
    # Get all Excel files from data folder
    excel_files = [f for f in os.listdir(data_folder) if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print(f"No Excel files found in '{data_folder}' folder!")
        return
    
    print(f"Found {len(excel_files)} file(s) to process")
    
    # Dictionary to store all data organized by event
    men_data = {}
    women_data = {}
    men_age_groups = []
    women_age_groups = []
    men_event_order = []
    women_event_order = []
    
    # Process each file
    for filename in sorted(excel_files):
        filepath = os.path.join(data_folder, filename)
        course, sex, age_start, age_end = parse_filename(filename)
        
        if not course or not sex:
            print(f"Skipping {filename} - doesn't match naming convention")
            continue
        
        age_group = age_end  # Use YY (end age) as specified
        age_group_label = f"{course}_{sex}_{age_group}"
        
        # Separate men and women data
        if sex == "Men":
            if age_group_label not in men_age_groups:
                men_age_groups.append(age_group_label)
            age_groups_list = men_age_groups
            data_dict = men_data
            event_order_list = men_event_order
        else:  # Women
            if age_group_label not in women_age_groups:
                women_age_groups.append(age_group_label)
            age_groups_list = women_age_groups
            data_dict = women_data
            event_order_list = women_event_order
        
        # Process the file
        file_results = process_file(filepath)
        
        # Store results organized by event, maintaining order
        for result in file_results:
            event_name = result['event']
            
            if event_name not in data_dict:
                data_dict[event_name] = {}
                event_order_list.append(event_name)  # Track order as events appear
            
            data_dict[event_name][age_group_label] = {
                '50th_time': result['50th_time'],
                'adjusted_times': result['adjusted_times']
            }
    
    # Create output DataFrames
    if men_data or women_data:
        output_file = 'swim_rankings_50th_place_summary.xlsx'
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Process Men's data
                if men_data:
                    men_rows = []
                
                    for event_name in men_event_order:
                        row = {'Event': event_name}
                    
                        for age_group_label in men_age_groups:
                            # FIXED: Check if age_group exists for this event
                            if age_group_label in men_data[event_name]:
                                row[f'{age_group_label}_50th'] = men_data[event_name][age_group_label]['50th_time']
                            
                                # Add all percentage calculations
                                for pct in [10, 11, 11.5, 12, 12.5]:
                                    pct_label = str(pct).replace('.', '_')
                                    row[f'{age_group_label}_+{pct_label}%'] = men_data[event_name][age_group_label]['adjusted_times'][pct]['time']
                                    row[f'{age_group_label}_+{pct_label}%_Rank'] = men_data[event_name][age_group_label]['adjusted_times'][pct]['rank']
                            else:
                                # FIXED: This else now properly aligned with the if
                                row[f'{age_group_label}_50th'] = ''
                                for pct in [10, 11, 11.5, 12, 12.5]:
                                    pct_label = str(pct).replace('.', '_')
                                    row[f'{age_group_label}_+{pct_label}%'] = ''
                                    row[f'{age_group_label}_+{pct_label}%_Rank'] = ''
                    
                        men_rows.append(row)
                
                    men_df = pd.DataFrame(men_rows)
                
                    # Reorder columns
                    men_cols = ['Event']
                    for age_group_label in men_age_groups:
                        men_cols.append(f'{age_group_label}_50th')
                        for pct in [10, 11, 11.5, 12, 12.5]:
                            pct_label = str(pct).replace('.', '_')
                            men_cols.append(f'{age_group_label}_+{pct_label}%')
                            men_cols.append(f'{age_group_label}_+{pct_label}%_Rank')
                
                    men_df = men_df[men_cols]
                    men_df.to_excel(writer, index=False, sheet_name='Men')
            
                # Process Women's data
                if women_data:
                    women_rows = []
                
                    for event_name in women_event_order:
                        row = {'Event': event_name}
                    
                        for age_group_label in women_age_groups:
                            # FIXED: Check if age_group exists for this event
                            if age_group_label in women_data[event_name]:
                                row[f'{age_group_label}_50th'] = women_data[event_name][age_group_label]['50th_time']
                            
                                # Add all percentage calculations
                                for pct in [10, 11, 11.5, 12, 12.5]:
                                    pct_label = str(pct).replace('.', '_')
                                    row[f'{age_group_label}_+{pct_label}%'] = women_data[event_name][age_group_label]['adjusted_times'][pct]['time']
                                    row[f'{age_group_label}_+{pct_label}%_Rank'] = women_data[event_name][age_group_label]['adjusted_times'][pct]['rank']
                            else:
                                # FIXED: This else now properly aligned with the if
                                row[f'{age_group_label}_50th'] = ''
                                for pct in [10, 11, 11.5, 12, 12.5]:
                                    pct_label = str(pct).replace('.', '_')
                                    row[f'{age_group_label}_+{pct_label}%'] = ''
                                    row[f'{age_group_label}_+{pct_label}%_Rank'] = ''
                    
                        women_rows.append(row)
                
                    women_df = pd.DataFrame(women_rows)
                
                    # Reorder columns
                    women_cols = ['Event']
                    for age_group_label in women_age_groups:
                        women_cols.append(f'{age_group_label}_50th')
                        for pct in [10, 11, 11.5, 12, 12.5]:
                            pct_label = str(pct).replace('.', '_')
                            women_cols.append(f'{age_group_label}_+{pct_label}%')
                            women_cols.append(f'{age_group_label}_+{pct_label}%_Rank')
                
                    women_df = women_df[women_cols]
                    women_df.to_excel(writer, index=False, sheet_name='Women')
        
            print(f"\n{'='*60}")
            print(f"Results exported to: {output_file}")
            if men_data:
                print(f"Men's events: {len(men_data)}")
                print(f"Men's age groups: {len(men_age_groups)}")
            if women_data:
                print(f"Women's events: {len(women_data)}")
                print(f"Women's age groups: {len(women_age_groups)}")
            print(f"{'='*60}")
        
            # Create simplified output file
            create_simplified_output(men_data, women_data, men_age_groups, women_age_groups, 
                                    men_event_order, women_event_order)
        
        except PermissionError:
            print(f"\n{'='*60}")
            print(f"ERROR: Cannot write to '{output_file}'")
            print(f"Please close the file if it's open in Excel and try again.")
            print(f"{'='*60}")
            return
    else:
        print("\nNo 50th place times found in any files!")

if __name__ == "__main__":
    main()