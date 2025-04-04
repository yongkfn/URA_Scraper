import requests
import os
import sys
import pandas as pd
import glob
from datetime import datetime
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Define constants
DOWNLOAD_DIR = "ura_land_sales_data"
REPORTS_DIR = "comparison_reports"
URL = "https://www.ura.gov.sg/-/media/Corporate/Land-Sales/Past-Sales-Sites/ura-vacant-sites.xlsx"
FORM_URL = "https://form.gov.sg/67ee3684e8ee3d46f1938396"

# ===== CUSTOM XPATHS FOR FORM FIELDS =====
# Fill in the exact XPaths for each form field below
# You can obtain these using browser developer tools (F12) -> right-click element -> Copy -> Copy XPath
FORM_XPATHS = {
    # Date fields
    "date_of_launch": '//*[@id="67ee3e1c70f6b1adf869c0ef"]',  # XPath for the Date of Launch input field
    "date_of_tender_closing": '//*[@id="67ee3e3c8f9e0a097696ff56"]',  # XPath for the Date of Tender Closing input field
    "date_of_award": '//*[@id="67ee370ab6c1d008404f165f"]',  # XPath for the Date of Award input field
    "location": '//*[@id="67ee370ab6c1d008404f1661"]',  # XPath for the Location input field
    "type_of_development": '//*[@id="67ee43a970f6b1adf86b6cac"]',  # XPath for the Type of Development Allowed dropdown
    # Text input fields

    "lease_years": '//*[@id="67ee370ab6c1d008404f1663"]',  # XPath for the Lease (years) input field
    "dev_code": '//*[@id="67ee370ab6c1d008404f1666"]',  # XPath for the Type of Devt Code input field
    "site_area": '//*[@id="67ee370ab6c1d008404f1667"]',  # XPath for the Site Area (m2) input field
    "name_of_successful_tenderer": '//*[@id="67ee370ab6c1d008404f166f"]',
    "successful_tender_price": '//*[@id="67ee4e242f9579d157d348b7"]',
    
    
    # Submit button
    "submit_button": '//*[@id="root"]/div/div[2]/div[4]/div[1]/div[1]/form/div[2]/button'  # XPath for the Submit button
}
# =========================================

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

def format_date_for_form(date_value):
    """Convert various date formats to dd/mm/yyyy for form input"""
    if pd.isna(date_value):
        return ""
        
    # If it's already a string in a recognizable format, parse it
    if isinstance(date_value, str):
        # Try various formats
        formats = ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%m-%d-%Y']
        for fmt in formats:
            try:
                date_obj = datetime.strptime(date_value, fmt)
                return date_obj.strftime('%d/%m/%Y')
            except ValueError:
                continue
                
        # If it contains a time component, try to parse that
        formats_with_time = ['%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S']
        for fmt in formats_with_time:
            try:
                date_obj = datetime.strptime(date_value, fmt)
                return date_obj.strftime('%d/%m/%Y')
            except ValueError:
                continue
                
    # If it's a pandas Timestamp or datetime object
    elif isinstance(date_value, (pd.Timestamp, datetime)):
        return date_value.strftime('%d/%m/%Y')
    
    # Fall back to string representation and handle manually
    print(f"Warning: Could not format date {date_value}, please enter manually")
    return str(date_value)

def submit_to_form_with_xpaths(entry_data, xpaths):
    """Submit data to the form using Selenium with custom XPaths"""
    print(f"\n--- SUBMITTING TO FORM WITH CUSTOM XPATHS ---")
    
    # Set up Chrome options
    chrome_options = Options()
    # Uncomment the line below to run headless (without UI)
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920,1080")
    
    try:
        # Initialize the driver
        print("Initializing browser...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        # Open the form
        print(f"Opening form URL: {FORM_URL}")
        driver.get(FORM_URL)
        
        # Wait for the form to load
        wait = WebDriverWait(driver, 30)
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))
            print("Form loaded successfully")
        except TimeoutException:
            print("Warning: Timeout waiting for form to load. Continuing anyway...")
        
        time.sleep(3)  # Give extra time for form to fully render
        
        # Get a screenshot of the form before filling
        before_screenshot_path = os.path.join(REPORTS_DIR, f"form_before_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        driver.save_screenshot(before_screenshot_path)
        print(f"Initial form screenshot saved to: {before_screenshot_path}")
        
        # Print column names from Excel for reference
        print("\nAvailable data in Excel row:")
        for col, val in entry_data.items():
            print(f"  {col}: {val}")
        
        # Map of form fields to data columns and XPaths
        field_mappings = [
            {
                "field_name": "Date of Launch",
                "data_column": "Date of Launch",
                "xpath": xpaths.get("date_of_launch", ""),
                "is_date": True
            },
            {
                "field_name": "Date of Tender Closing",
                "data_column": "Date of Tender Closing",
                "xpath": xpaths.get("date_of_tender_closing", ""),
                "is_date": True
            },
            {
                "field_name": "Date of Award",
                "data_column": "Date of Award",
                "xpath": xpaths.get("date_of_award", ""),
                "is_date": True
            },
            {
                "field_name": "Location",
                "data_column": "Location",
                "xpath": xpaths.get("location", ""),
                "is_date": False
            },
            {
                "field_name": "Type of Development Allowed",
                "data_column": "Type of Development Allowed",
                "xpath": xpaths.get("type_of_development", ""),
            },
            {
                "field_name": "Lease (years)",
                "data_column": "Lease (years)",
                "xpath": xpaths.get("lease_years", ""),
                "is_date": False
            },
            {
                "field_name": "Type of Devt Code",
                "data_column": ["Type of Devt Code", "Type of Devt Code (Click here for explanation)"],
                "xpath": xpaths.get("dev_code", ""),
                "is_date": False
            },
            {
                "field_name": "Site Area (m2)",
                "data_column": "Site Area (m2)",
                "xpath": xpaths.get("site_area", ""),
                "is_date": False
            },
            {
                "field_name": "Name of Successful Tenderer",
                "data_column": "Name of Successful Tenderer",
                "xpath": xpaths.get("name_of_successful_tenderer", ""),
                "is_date": False
            },
            {
                "field_name": "Successful Tender Price ",
                "data_column": "Successful Tender Price ",
                "xpath": xpaths.get("successful_tender_price", ""),
                "is_date": False
            },
        ] 
        
        # Process each field
        for field in field_mappings:
            try:
                print(f"\nFilling {field['field_name']}...")
                
                # Skip if xpath is empty
                if not field['xpath']:
                    print(f"  Skipping {field['field_name']} - no XPath provided")
                    continue
                
                # Handle multiple possible column names
                data_value = None
                if isinstance(field['data_column'], list):
                    for col in field['data_column']:
                        if col in entry_data and not pd.isna(entry_data[col]):
                            data_value = entry_data[col]
                            break
                else:
                    if field['data_column'] in entry_data and not pd.isna(entry_data[field['data_column']]):
                        data_value = entry_data[field['data_column']]
                
                if data_value is None:
                    print(f"  No data for {field['field_name']}")
                    continue
                
                # Format date fields
                if field.get('is_date', False):
                    data_value = format_date_for_form(data_value)
                else:
                    data_value = str(data_value)
                
                # Handle dropdown fields
                if field.get('is_dropdown', False):
                    try:
                        # Click to open dropdown
                        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, field['xpath'])))
                        dropdown.click()
                        time.sleep(1)
                        
                        # If dropdown option xpath is provided
                        if field.get('dropdown_option_xpath'):
                            # Find option with custom xpath
                            option_xpath = field['dropdown_option_xpath'].replace('{value}', data_value)
                            option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
                            option.click()
                        else:
                            # Try to find option by text
                            option = wait.until(EC.element_to_be_clickable(
                                (By.XPATH, f"//div[contains(@class, 'Select__option') and contains(text(), '{data_value}')]")))
                            option.click()
                        
                        print(f"  Selected {field['field_name']}: {data_value}")
                    except Exception as e:
                        print(f"  Error selecting {field['field_name']}: {str(e)}")
                        # Try to click Escape key to close dropdown if it's still open
                        try:
                            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        except:
                            pass
                else:
                    # Regular input field
                    try:
                        input_element = wait.until(EC.presence_of_element_located((By.XPATH, field['xpath'])))
                        input_element.clear()
                        input_element.click()
                        input_element.send_keys(data_value)
                        
                        if field.get('is_date', False):
                            input_element.send_keys(Keys.TAB)  # Tab out to confirm date
                        
                        print(f"  Filled {field['field_name']}: {data_value}")
                    except Exception as e:
                        print(f"  Error filling {field['field_name']}: {str(e)}")
            except Exception as e:
                print(f"  Error processing {field['field_name']}: {str(e)}")
        
        # Take a screenshot of the filled form
        screenshot_path = os.path.join(REPORTS_DIR, f"form_filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        driver.save_screenshot(screenshot_path)
        print(f"\nForm screenshot saved to: {screenshot_path}")
        
        # Ask user to verify and submit
        print("\nForm filled with available data. Please verify the information in the browser.")
        submit_choice = input("Submit the form? (y/n): ")
        
        if submit_choice.lower() == 'y':
            # Find and click the submit button
            try:
                if xpaths.get("submit_button", ""):
                    # Use provided submit button xpath
                    submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, xpaths["submit_button"])))
                    submit_button.click()
                else:
                    # Try different approaches to find the submit button
                    try:
                        submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Submit')]")))
                    except:
                        try:
                            submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
                        except:
                            submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.submit-button")))
                    
                    submit_button.click()
                
                print("Clicked submit button")
                
                # Wait for confirmation page
                try:
                    wait.until(EC.url_changes(FORM_URL))
                    print("Form submitted successfully!")
                except:
                    print("Form may have been submitted, but confirmation couldn't be verified")
            except Exception as e:
                print(f"Error submitting form: {str(e)}")
        else:
            print("Form submission cancelled.")
        
        # Wait before closing the browser
        time.sleep(5)
        driver.quit()
        
    except Exception as e:
        print(f"Error during form submission process: {str(e)}")
        try:
            # Take a screenshot if there's an error
            screenshot_path = os.path.join(REPORTS_DIR, f"form_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            driver.save_screenshot(screenshot_path)
            print(f"Error screenshot saved to {screenshot_path}")
            driver.quit()
        except:
            print("Could not save error screenshot")

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
    
    # Create report and submit to form if new entries are found
    if len(new_entries) > 0:
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
            new_entries.to_excel(writer, sheet_name='New Entries', index=False)
        
        print(f"\nDetailed report saved to: {report_path}")
        
        # Process each new entry for form submission
        for i, (_, entry) in enumerate(new_entries.iterrows(), 1):
            print(f"\nProcessing new entry {i} of {len(new_entries)}")
            
            # Display entry details
            print("Entry details:")
            for col, val in entry.items():
                print(f"  {col}: {val}")
            
            # Ask if user wants to submit this entry
            submit_choice = input(f"\nDo you want to submit entry {i} to the form? (y/n): ")
            
            if submit_choice.lower() == 'y':
                # Convert the pandas Series to a dictionary for easier handling
                entry_dict = entry.to_dict()
                
                # Submit to form
                submit_to_form_with_xpaths(entry_dict, FORM_XPATHS)
                
                # Wait a bit before processing next entry
                time.sleep(2)
            else:
                print("Skipping this entry.")
    else:
        print("\nNo new entries found between the files.")

def main():
    """Main function to download and compare Excel files"""
    print("URA Vacant Sites Download, Comparison, and Form Submission Tool")
    print("=============================================================")
    
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