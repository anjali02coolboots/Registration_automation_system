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

# Now append to Registration_Template.xlsx
print(f"\nAppending data to {template_path}...")

if os.path.exists(template_path):
    # File exists, append data below existing data
    print("Template file exists. Appending new data...")
    
    # Read existing data
    existing_df = pd.read_excel(template_path, engine='openpyxl')
    
    # Combine existing data with new data
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    
    # Save combined data
    combined_df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"Appended {len(df)} rows to existing {len(existing_df)} rows")
    print(f"Total rows in template: {len(combined_df)}")
    
else:
    # File doesn't exist, create new file with current data
    print("Template file doesn't exist. Creating new file...")
    df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"Created new template file with {len(df)} rows")

print("\nProcessing complete!")
print(f"Date processed: {yesterday_str}")
print(f"Intermediate output saved to: {output_path}")
print(f"Final template saved to: {template_path}")
print("\nFirst few rows of today's processed data:")
print(df.head())