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
            '0-30': (delta_days, 0 <= delta_days <= 30),
            '30-60': (delta_days, 31 <= delta_days <= 60),
            '60-90': (delta_days, 61 <= delta_days <= 90),
            '>90': (delta_days, delta_days > 90)
        }
    return {'0-30': (None, False), '30-60': (None, False), '60-90': (None, False), '>90': (None, False)}

# Initialize columns to store differences
diff_columns = [
    'due date diff 0-30 days', 'due date diff 30-60 days', 'due date diff 60-90 days', 'due date diff >90 days',
    'exception expiration diff 0-30 days', 'exception expiration diff 30-60 days', 'exception expiration diff 60-90 days', 'exception expiration diff >90 days'
]

for column in diff_columns:
    extracted_data[column] = None

# Calculate date ranges for Due Date
for index, row in extracted_data.iterrows():
    due_date_ranges = check_date_ranges(row['Due Date'], current_date)
    extracted_data.at[index, 'due date diff 0-30 days'] = due_date_ranges['0-30'][0]
    extracted_data.at[index, 'due date 0-30 days'] = due_date_ranges['0-30'][1]
    extracted_data.at[index, 'due date diff 30-60 days'] = due_date_ranges['30-60'][0]
    extracted_data.at[index, 'due date 30-60 days'] = due_date_ranges['30-60'][1]
    extracted_data.at[index, 'due date diff 60-90 days'] = due_date_ranges['60-90'][0]
    extracted_data.at[index, 'due date 60-90 days'] = due_date_ranges['60-90'][1]
    extracted_data.at[index, 'due date diff >90 days'] = due_date_ranges['>90'][0]
    extracted_data.at[index, 'due date >90 days'] = due_date_ranges['>90'][1]

# Calculate date ranges for Exception Expiration
for index, row in extracted_data.iterrows():
    exception_expiration_ranges = check_date_ranges(row['Exception Expiration'], current_date)
    extracted_data.at[index, 'exception expiration diff 0-30 days'] = exception_expiration_ranges['0-30'][0]
    extracted_data.at[index, 'exception expiration 0-30 days'] = exception_expiration_ranges['0-30'][1]
    extracted_data.at[index, 'exception expiration diff 30-60 days'] = exception_expiration_ranges['30-60'][0]
    extracted_data.at[index, 'exception expiration 30-60 days'] = exception_expiration_ranges['30-60'][1]
    extracted_data.at[index, 'exception expiration diff 60-90 days'] = exception_expiration_ranges['60-90'][0]
    extracted_data.at[index, 'exception expiration 60-90 days'] = exception_expiration_ranges['60-90'][1]
    extracted_data.at[index, 'exception expiration diff >90 days'] = exception_expiration_ranges['>90'][0]
    extracted_data.at[index, 'exception expiration >90 days'] = exception_expiration_ranges['>90'][1]

# Save the modified DataFrame to a new Excel file
output_path = 'date_ranges_vulnerability_data.xlsx'  # Change this to your desired output path
extracted_data.to_excel(output_path, index=False)

print("Date ranges analysis saved to", output_path)
