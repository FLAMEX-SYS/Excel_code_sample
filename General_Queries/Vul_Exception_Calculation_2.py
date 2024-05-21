import pandas as pd
from datetime import datetime

# Load the vulnerability data
file_path = 'path_to_your_vulnerability_data.xlsx'  # Change this to the path of your vulnerability data file
vul_data = pd.read_excel(file_path)

# Extract the relevant columns
extracted_columns = ['hostname', 'Due Date', 'Status', 'CVSS', 'Exception Expiration']
extracted_data = vul_data[extracted_columns]

# Ensure that date columns are parsed correctly and are timezone-naive
extracted_data['Due Date'] = pd.to_datetime(extracted_data['Due Date'], errors='coerce').dt.tz_localize(None)
extracted_data['Exception Expiration'] = pd.to_datetime(extracted_data['Exception Expiration'], errors='coerce').dt.tz_localize(None)

# Print the first value under the "Due Date" column
first_due_date = extracted_data['Due Date'].iloc[0]
print("First Due Date:", first_due_date)

# Define the current date and ensure it is timezone-naive
current_date = pd.Timestamp('today').normalize()

# Initialize new columns for date range checks with False values
date_range_columns = [
    'due date 0-30 days', 'due date 30-60 days', 'due date 60-90 days', 'due date >90 days',
    'exception expiration 0-30 days', 'exception expiration 30-60 days', 'exception expiration 60-90 days', 'exception expiration >90 days'
]

for column in date_range_columns:
    extracted_data[column] = False

# Function to check date ranges and get the difference in days
def check_date_ranges(date_value, current_date):
    if pd.notnull(date_value):
        delta_days = (date_value - current_date).days  # Get the number of days as an integer
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

# Calculate date ranges for Exception Expiration, setting blank values to False by default
for index, row in extracted_data.iterrows():
    if pd.isnull(row['Exception Expiration']):
        extracted_data.at[index, 'exception expiration 0-30 days'] = False
        extracted_data.at[index, 'exception expiration 30-60 days'] = False
        extracted_data.at[index, 'exception expiration 60-90 days'] = False
        extracted_data.at[index, 'exception expiration >90 days'] = False
    else:
        exception_expiration_ranges = check_date_ranges(row['Exception Expiration'], current_date)
        extracted_data.at[index, 'exception expiration 0-30 days'] = exception_expiration_ranges['0-30']
        extracted_data.at[index, 'exception expiration 30-60 days'] = exception_expiration_ranges['30-60']
        extracted_data.at[index, 'exception expiration 60-90 days'] = exception_expiration_ranges['60-90']
        extracted_data.at[index, 'exception expiration >90 days'] = exception_expiration_ranges['>90']

# Group by hostname and count the number of True values in each column
aggregated_data = extracted_data.groupby('hostname').agg({
    'Due Date': 'count',
    'due date 0-30 days': 'sum',
    'due date 30-60 days': 'sum',
    'due date 60-90 days': 'sum',
    'due date >90 days': 'sum',
    'exception expiration 0-30 days': 'sum',
    'exception expiration 30-60 days': 'sum',
    'exception expiration 60-90 days': 'sum',
    'exception expiration >90 days': 'sum'
}).reset_index()

# Rename the columns for clarity
aggregated_data = aggregated_data.rename(columns={
    'Due Date': 'vulnerability_count',
    'due date 0-30 days': 'due date 0-30 days count',
    'due date 30-60 days': 'due date 30-60 days count',
    'due date 60-90 days': 'due date 60-90 days count',
    'due date >90 days': 'due date >90 days count',
    'exception expiration 0-30 days': 'exception expiration 0-30 days count',
    'exception expiration 30-60 days': 'exception expiration 30-60 days count',
    'exception expiration 60-90 days': 'exception expiration 60-90 days count',
    'exception expiration >90 days': 'exception expiration >90 days count',
})

# Save the aggregated DataFrame to a new Excel file
output_path_aggregated = 'aggregated_vulnerability_data.xlsx'  # Change this to your desired output path
aggregated_data.to_excel(output_path_aggregated, index=False)

print("Aggregated data saved to", output_path_aggregated)
