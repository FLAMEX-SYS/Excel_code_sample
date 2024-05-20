import pandas as pd
from datetime import datetime

# Load the vulnerability data
file_path = 'path_to_your_vulnerability_data.xlsx'  # Change this to the path of your vulnerability data file
vul_data = pd.read_excel(file_path)

# Extract the relevant columns
extracted_columns = ['hostname', 'Due Date', 'Status', 'CVSS', 'Exception Expiration']
extracted_data = vul_data[extracted_columns]

# Ensure that date columns are parsed correctly
extracted_data['Due Date'] = pd.to_datetime(extracted_data['Due Date'], errors='coerce')
extracted_data['Exception Expiration'] = pd.to_datetime(extracted_data['Exception Expiration'], errors='coerce')

# Define the current date
current_date = pd.Timestamp('today')

# Initialize new columns for date range checks with False values
date_range_columns = [
    'due date 0-30 days', 'due date 30-60 days', 'due date 60-90 days', 'due date >90 days',
    'exception expiration 0-30 days', 'exception expiration 30-60 days', 'exception expiration 60-90 days', 'exception expiration >90 days'
]

for column in date_range_columns:
    extracted_data[column] = False

# Function to check date ranges
def check_date_ranges(date_value, current_date):
    if pd.notnull(date_value):
        delta_days = (date_value - current_date).days
        return {
            '0-30': 0 <= delta_days <= 30,
            '30-60': 31 <= delta_days <= 60,
            '60-90': 61 <= delta_days <= 90,
            '>90': delta_days > 90
        }
    return {'0-30': False, '30-60': False, '60-90': False, '>90': False}

# Calculate date ranges for Due Date
for index, row in extracted_data.iterrows():
    due_date_ranges = check_date_ranges(row['Due Date'], current_date)
    extracted_data.at[index, 'due date 0-30 days'] = due_date_ranges['0-30']
    extracted_data.at[index, 'due date 30-60 days'] = due_date_ranges['30-60']
    extracted_data.at[index, 'due date 60-90 days'] = due_date_ranges['60-90']
    extracted_data.at[index, 'due date >90 days'] = due_date_ranges['>90']

# Calculate date ranges for Exception Expiration
for index, row in extracted_data.iterrows():
    exception_expiration_ranges = check_date_ranges(row['Exception Expiration'], current_date)
    extracted_data.at[index, 'exception expiration 0-30 days'] = exception_expiration_ranges['0-30']
    extracted_data.at[index, 'exception expiration 30-60 days'] = exception_expiration_ranges['30-60']
    extracted_data.at[index, 'exception expiration 60-90 days'] = exception_expiration_ranges['60-90']
    extracted_data.at[index, 'exception expiration >90 days'] = exception_expiration_ranges['>90']

# Save the modified DataFrame to a new Excel file
output_path = 'date_ranges_vulnerability_data.xlsx'  # Change this to your desired output path
extracted_data.to_excel(output_path, index=False)

print("Date ranges analysis saved to", output_path)
