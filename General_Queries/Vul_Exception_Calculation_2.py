To aggregate the data by `hostname` and compute the counts for the different due date ranges, you can use the `groupby` function in pandas. This will allow you to summarize the data for each hostname and count the number of vulnerabilities falling into each specified date range.

Here’s how you can modify the script to achieve this:

### Updated Python Script

```python
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

# Define the current date and ensure it is timezone-naive
current_date = pd.Timestamp('today').normalize()

# Initialize new columns for date range checks with False values and for storing the differences in days
date_range_columns = [
    'due date diff 0-30 days', 'due date diff 30-60 days', 'due date diff 60-90 days', 'due date diff >90 days',
    'due date 0-30 days', 'due date 30-60 days', 'due date 60-90 days', 'due date >90 days',
    'exception expiration diff 0-30 days', 'exception expiration diff 30-60 days', 'exception expiration diff 60-90 days', 'exception expiration diff >90 days',
    'exception expiration 0-30 days', 'exception expiration 30-60 days', 'exception expiration 60-90 days', 'exception expiration >90 days'
]

for column in date_range_columns:
    extracted_data[column] = 0

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

# Group by hostname and aggregate the counts
aggregated_data = extracted_data.groupby('hostname').agg({
    'hostname': 'first',
    'due date diff 0-30 days': 'count',
    'due date 0-30 days': 'sum',
    'due date diff 30-60 days': 'sum',
    'due date 30-60 days': 'sum',
    'due date diff 60-90 days': 'sum',
    'due date 60-90 days': 'sum',
    'due date diff >90 days': 'sum',
    'due date >90 days': 'sum',
    'exception expiration diff 0-30 days': 'sum',
    'exception expiration 0-30 days': 'sum',
    'exception expiration diff 30-60 days': 'sum',
    'exception expiration 30-60 days': 'sum',
    'exception expiration diff 60-90 days': 'sum',
    'exception expiration 60-90 days': 'sum',
    'exception expiration diff >90 days': 'sum',
    'exception expiration >90 days': 'sum'
}).reset_index()

# Rename the columns for clarity
aggregated_data = aggregated_data.rename(columns={
    'due date diff 0-30 days': 'due date 0-30 days count',
    'due date diff 30-60 days': 'due date 30-60 days count',
    'due date diff 60-90 days': 'due date 60-90 days count',
    'due date diff >90 days': 'due date >90 days count',
    'exception expiration diff 0-30 days': 'exception expiration 0-30 days count',
    'exception expiration diff 30-60 days': 'exception expiration 30-60 days count',
    'exception expiration diff 60-90 days': 'exception expiration 60-90 days count',
    'exception expiration diff >90 days': 'exception expiration >90 days count',
})

# Save the aggregated DataFrame to a new Excel file
output_path_aggregated = 'aggregated_vulnerability_data.xlsx'  # Change this to your desired output path
aggregated_data.to_excel(output_path_aggregated, index=False)

print("Aggregated data saved to", output_path_aggregated)
