import requests
import os
import sys
import pandas as pd
import glob
from datetime import datetime
import re
import time

# Define constants
DOWNLOAD_DIR = "ura_land_sales_data"
REPORTS_DIR = "comparison_reports"
URL = "https://www.ura.gov.sg/-/media/Corporate/Land-Sales/Past-Sales-Sites/ura-vacant-sites.xlsx"

def ensure_directories_exist():
    """Create necessary directories if they don't exist"""
    for directory in [DOWNLOAD_DIR, REPORTS_DIR]:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Created directory: {directory}")

def download_vacant_sites_file():
    """Download the URA vacant sites Excel file"""
    print("\n--- DOWNLOADING LATEST FILE ---")
    print("Starting URA vacant sites Excel download...")
    
    # Create directory if it doesn't exist
    ensure_directories_exist()
    
    # Set headers to mimic a browser
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    # Create filename with date prefix
    date_prefix = datetime.now().strftime('%Y%m%d')
    filename = f"{date_prefix}_ura-vacant-sites.xlsx"
    output_path = os.path.join(DOWNLOAD_DIR, filename)
    
    # Check if today's file already exists
    if os.path.exists(output_path):
        print(f"Today's file already exists: {output_path}")
        return output_path
    
    try:
        # Download the file
        print(f"Downloading from: {URL}")
        response = requests.get(URL, headers=headers)
        response.raise_for_status()  # Check for HTTP errors
        
        # Save to a temporary file first
        temp_path = output_path + ".tmp"
        with open(temp_path, 'wb') as file:
            file.write(response.content)
        
        # If download succeeded, rename to final name
        os.rename(temp_path, output_path)
        print(f"File saved successfully to: {output_path}")
        
        return output_path
            
    except Exception as e:
        print(f"Error during download: {str(e)}")
        return None

def find_previous_excel_file(current_file):
    """Find the most recent Excel file before the current one"""
    # Get all Excel files matching the pattern
    pattern = "*ura-vacant-sites*.xlsx"
    excel_files = glob.glob(os.path.join(DOWNLOAD_DIR, pattern))
    
    if not excel_files or len(excel_files) <= 1:
        print("No previous Excel files found for comparison")
        return None
    
    # Extract date from filename (assuming format: YYYYMMDD_filename.xlsx)
    def extract_date(filename):
        basename = os.path.basename(filename)
        match = re.search(r'(\d{8})_', basename)
        if match:
            date_str = match.group(1)
            try:
                return datetime.strptime(date_str, '%Y%m%d')
            except ValueError:
                return datetime.min
        return datetime.min
    
    # Sort files by date (newest first)
    sorted_files = sorted(excel_files, key=extract_date, reverse=True)
    
    # Find the current file in the list
    try:
        current_index = sorted_files.index(current_file)
    except ValueError:
        # Current file not found in list, just return the most recent file
        return sorted_files[0] if current_file != sorted_files[0] else sorted_files[1] if len(sorted_files) > 1 else None
    
    # Return the next most recent file
    if current_index < len(sorted_files) - 1:
        return sorted_files[current_index + 1]
    else:
        print("No previous file found")
        return None

def identify_key_columns(df):
    """Try to identify columns that could serve as unique identifiers"""
    # Common naming patterns for ID columns
    id_patterns = ['id', 'key', 'code', 'no', 'number', 'lot', 'name', 'location']
    
    potential_keys = []
    
    # First, look for columns with names suggesting they're IDs
    for col in df.columns:
        for pattern in id_patterns:
            if pattern.lower() in col.lower():
                potential_keys.append(col)
                break
    
    # If we found potential key columns, check if they're unique
    if potential_keys:
        # Check if any single column is unique
        for col in potential_keys:
            if df[col].nunique() == len(df):
                return [col]  # Found a single column that's unique
        
        # Check if combinations of columns are unique
        for i in range(2, len(potential_keys) + 1):
            for combo in pd.Series(potential_keys).drop_duplicates().tolist():
                if df.duplicated(subset=combo).sum() == 0:
                    return combo  # Found a combination that's unique
    
    # As a fallback, try all columns to see if any are unique
    for col in df.columns:
        if df[col].nunique() == len(df):
            return [col]
    
    # If we get here, no single column is unique, try combinations
    print("Warning: No unique identifier columns found. Comparison may not be accurate.")
    return list(df.columns)  # Use all columns as a last resort

def compare_excel_files(newer_file, older_file):
    """Compare two Excel files and identify new entries"""
    print("\n--- COMPARING FILES ---")
    print(f"Comparing Excel files:")
    print(f"  Newer file: {os.path.basename(newer_file)}")
    print(f"  Older file: {os.path.basename(older_file)}")
    
    # Load Excel files into pandas DataFrames
    try:
        df_newer = pd.read_excel(newer_file)
        df_older = pd.read_excel(older_file)
    except Exception as e:
        print(f"Error loading Excel files: {str(e)}")
        return
    
    # Get basic information about the files
    print("\nBasic Information:")
    print(f"  Newer file: {len(df_newer)} rows, {len(df_newer.columns)} columns")
    print(f"  Older file: {len(df_older)} rows, {len(df_older.columns)} columns")
    
    # Check if column names match
    if set(df_newer.columns) != set(df_older.columns):
        print("\nColumn differences detected:")
        only_in_newer = set(df_newer.columns) - set(df_older.columns)
        only_in_older = set(df_older.columns) - set(df_newer.columns)
        
        if only_in_newer:
            print(f"  Columns only in newer file: {', '.join(only_in_newer)}")
        if only_in_older:
            print(f"  Columns only in older file: {', '.join(only_in_older)}")
        
        # Use only common columns for comparison
        common_columns = list(set(df_newer.columns) & set(df_older.columns))
        df_newer = df_newer[common_columns]
        df_older = df_older[common_columns]
        
        print(f"  Using {len(common_columns)} common columns for comparison")
    
    # Identify key columns for comparison
    key_columns = identify_key_columns(df_newer)
    print(f"\nUsing the following columns as unique identifiers: {', '.join(key_columns)}")
    
    # Set index for comparison
    df_newer_indexed = df_newer.set_index(key_columns)
    df_older_indexed = df_older.set_index(key_columns)
    
    # Find rows that are in newer but not in older file (new entries)
    new_entries_idx = df_newer_indexed.index.difference(df_older_indexed.index)
    new_entries = df_newer.loc[df_newer.set_index(key_columns).index.isin(new_entries_idx)]
    
    # Generate report
    print(f"\nComparison Results:")
    print(f"  New entries detected: {len(new_entries)}")
    
    # Create report filename with date and time
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_filename = f"new_entries_report_{timestamp}.xlsx"
    report_path = os.path.join(REPORTS_DIR, report_filename)
    
    # Create report only if new entries are found
    if len(new_entries) > 0:
        # Display some new entries in console
        print("\nSample of new entries:")
        sample_size = min(5, len(new_entries))
        for i, (_, row) in enumerate(new_entries.head(sample_size).iterrows(), 1):
            print(f"  Entry {i}:")
            for col, val in row.items():
                if col in key_columns:  # Highlight key columns
                    print(f"    *{col}: {val}")
                else:
                    print(f"    {col}: {val}")
            if i < sample_size:
                print()  # Empty line between entries
        
        # Create Excel report
        with pd.ExcelWriter(report_path) as writer:
            # Summary sheet
            summary_data = {
                'Metric': ['Date of comparison', 'Newer file', 'Older file', 
                           'Rows in newer file', 'Rows in older file', 'New entries found'],
                'Value': [datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 
                         os.path.basename(newer_file), os.path.basename(older_file),
                         len(df_newer), len(df_older), len(new_entries)]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
            
            # New entries sheet
            if not new_entries.empty:
                new_entries.to_excel(writer, sheet_name='New Entries', index=False)
        
        print(f"\nDetailed report saved to: {report_path}")
    else:
        print("\nNo new entries found between the files.")

def main():
    """Main function to download and compare Excel files"""
    print("URA Vacant Sites Download and New Entries Tracker")
    print("================================================")
    
    # Ensure directories exist
    ensure_directories_exist()
    
    # Download the latest file
    new_file = download_vacant_sites_file()
    
    if new_file:
        # Find the previous file for comparison
        previous_file = find_previous_excel_file(new_file)
        
        if previous_file:
            # Compare the files
            compare_excel_files(new_file, previous_file)
        else:
            print("\nNo previous file found for comparison.")
    else:
        print("Download failed. Cannot perform comparison.")

if __name__ == "__main__":
    main()