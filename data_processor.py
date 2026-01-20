import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook

# Get the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct file paths
csv_path = os.path.join(script_dir, 'exports', 'Registered_User_Source_Summary.csv')
lookup_path = os.path.join(script_dir, 'Source_TG_Latest.xlsx')
output_path = os.path.join(script_dir, 'exports', 'Processed_User_Summary.xlsx')
template_path = os.path.join(script_dir, 'Registration_Template.xlsx')

# Check if files exist
if not os.path.exists(csv_path):
    print(f"Error: CSV file not found at {csv_path}")
    exit(1)

if not os.path.exists(lookup_path):
    print(f"Error: Lookup file not found at {lookup_path}")
    exit(1)

# Read the CSV file with comma separator
print("Reading CSV file...")
df = pd.read_csv(csv_path)

# Get yesterday's date in the format 15-01-2026
yesterday = datetime.now() - timedelta(days=1)
yesterday_str = yesterday.strftime('%d-%m-%Y')

# Add Date column as the first column
df.insert(0, 'Date', yesterday_str)

# Keep only the required columns
columns_to_keep = ['Date', 'Registration Type', 'Registration Source', 'Campaign Source']
df = df[columns_to_keep]

# Read the lookup file (Category sheet)
print("Reading lookup file...")
lookup_df = pd.read_excel(lookup_path, sheet_name='Category')

# Perform XLOOKUP equivalent
print("Performing lookup...")

# Column B (index 1) - Source (Dashboard) - this is the lookup array
# Column C (index 2) - Actual Source - this is the return array
lookup_col_b = lookup_df.columns[1]  # Column B - Source (Dashboard)
lookup_col_c = lookup_df.columns[2]  # Column C - Actual Source

# Create lookup dictionary
lookup_dict = dict(zip(lookup_df[lookup_col_b], lookup_df[lookup_col_c]))

# Apply lookup to create New Source column
df['New Source'] = df['Registration Source'].map(lookup_dict)

# Fill NaN values with empty string
df['New Source'] = df['New Source'].fillna('')

# Save to Excel file (intermediate file)
print(f"Saving processed file to {output_path}...")
df.to_excel(output_path, index=False, engine='openpyxl')

# Now append to Registration_Template.xlsx with duplicate detection
print(f"\nAppending data to {template_path}...")

if os.path.exists(template_path):
    # File exists, check for duplicates before appending
    print("Template file exists. Checking for duplicates...")
    
    # Read existing data
    existing_df = pd.read_excel(template_path, engine='openpyxl')
    
    # Check if data for this date already exists
    existing_dates = existing_df['Date'].astype(str).unique() if 'Date' in existing_df.columns else []
    
    if yesterday_str in existing_dates:
        print(f"⚠️  WARNING: Data for {yesterday_str} already exists in template!")
        print(f"   Existing rows for this date: {len(existing_df[existing_df['Date'].astype(str) == yesterday_str])}")
        print(f"   New rows to add: {len(df)}")
        
        # Remove existing data for this date
        print(f"   Removing old data for {yesterday_str} and replacing with new data...")
        existing_df = existing_df[existing_df['Date'].astype(str) != yesterday_str]
        print(f"   ✓ Old data removed. Rows remaining: {len(existing_df)}")
    
    # Combine existing data with new data
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    
    # Sort by date (newest first)
    try:
        combined_df['Date_Parsed'] = pd.to_datetime(combined_df['Date'], format='%d-%m-%Y', errors='coerce')
        combined_df = combined_df.sort_values('Date_Parsed', ascending=False)
        combined_df = combined_df.drop(columns=['Date_Parsed'])
    except:
        print("   Note: Could not sort by date, keeping insertion order")
    
    # Save combined data
    combined_df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"✓ Appended {len(df)} rows")
    print(f"✓ Total rows in template: {len(combined_df)}")
    
else:
    # File doesn't exist, create new file with current data
    print("Template file doesn't exist. Creating new file...")
    df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"✓ Created new template file with {len(df)} rows")

print("\n" + "="*60)
print("PROCESSING COMPLETE!")
print("="*60)
print(f"Date processed: {yesterday_str}")
print(f"Intermediate output: {output_path}")
print(f"Final template: {template_path}")
print("\nFirst few rows of today's processed data:")
print(df.head())
print("="*60)