import pandas as pd
from datetime import datetime

# Load the first sheet with CM2, CM3, and asset ID
file_path_1 = 'path_to_your_first_sheet.xlsx'  # Change this to the path of your first sheet
cm_data = pd.read_excel(file_path_1)

# Load the second sheet with asset ID, vulnerability count, CVSS score, and other fields
file_path_2 = 'path_to_your_second_sheet.xlsx'  # Change this to the path of your second sheet
vul_data = pd.read_excel(file_path_2)

# Ensure the columns are named appropriately
# Assuming the columns in the first sheet are named 'asset_id', 'CM2', and 'CM3'
# Assuming the columns in the second sheet are named 'asset_id', 'vul_count', 'CVSS', 'Due Date', 'Exception Expiration', 'Status'
cm_data.columns = ['asset_id', 'CM2', 'CM3']  # Adjust if necessary
vul_data.columns = ['asset_id', 'vul_count', 'CVSS', 'Due Date', 'Exception Expiration', 'Status']  # Adjust if necessary

# Filter out rows where CM2 or CM3 are empty
cm_data_filtered = cm_data.dropna(subset=['CM2', 'CM3'])

# Ensure that date columns are parsed correctly and are timezone-naive
vul_data['Due Date'] = pd.to_datetime(vul_data['Due Date'], errors='coerce').dt.tz_localize(None)
vul_data['Exception Expiration'] = pd.to_datetime(vul_data['Exception Expiration'], errors='coerce').dt.tz_localize(None)

# Define the current date and ensure it is timezone-naive
current_date = pd.Timestamp('today').normalize()

# Initialize new columns for date range checks with False values
date_range_columns = [
    'due date 0-30 days', 'due date 30-60 days', 'due date 60-90 days', 'due date >90 days',
    'exception expiration 0-30 days', 'exception expiration 30-60 days', 'exception expiration 60-90 days', 'exception expiration >90 days'
]

for column in date_range_columns:
    vul_data[column] = False

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

# Calculate date ranges for Due Date only for high CVSS scores
for index, row in vul_data.iterrows():
    if row['CVSS'] > 7:
        due_date_ranges = check_date_ranges(row['Due Date'], current_date)
        vul_data.at[index, 'due date 0-30 days'] = due_date_ranges['0-30']
        vul_data.at[index, 'due date 30-60 days'] = due_date_ranges['30-60']
        vul_data.at[index, 'due date 60-90 days'] = due_date_ranges['60-90']
        vul_data.at[index, 'due date >90 days'] = due_date_ranges['>90']

        exception_expiration_ranges = check_date_ranges(row['Exception Expiration'], current_date)
        vul_data.at[index, 'exception expiration 0-30 days'] = exception_expiration_ranges['0-30']
        vul_data.at[index, 'exception expiration 30-60 days'] = exception_expiration_ranges['30-60']
        vul_data.at[index, 'exception expiration 60-90 days'] = exception_expiration_ranges['60-90']
        vul_data.at[index, 'exception expiration >90 days'] = exception_expiration_ranges['>90']

# Convert boolean columns to integers to avoid float issues
for column in date_range_columns:
    vul_data[column] = vul_data[column].astype(int)

# Add new columns for the different counts
vul_data['vul_count_high_cvss'] = vul_data['CVSS'].apply(lambda x: 1 if x > 7 else 0)
vul_data['vul_count_exceptions_only'] = vul_data['Exception Expiration'].apply(lambda x: 1 if pd.notnull(x) else 0)
vul_data['vul_count_closed'] = vul_data['Status'].apply(lambda x: 1 if x.lower() == 'closed' else 0)

# Merge the cm_data_filtered and vul_data DataFrames on the 'asset_id' column
merged_data = pd.merge(cm_data_filtered, vul_data, on='asset_id', how='inner')

# Group by CM2 and CM3 to calculate the required counts
consolidated_data = merged_data.groupby(['CM2', 'CM3']).agg(
    Asset_Count=('asset_id', 'nunique'),  # Count of unique asset IDs
    Vul_Count=('vul_count', 'sum'),  # Sum of vulnerability counts
    Vul_Count_High_CVSS=('vul_count_high_cvss', 'sum'),
    Vul_Count_Exceptions_Only=('vul_count_exceptions_only', 'sum'),
    Vul_Count_Exceptions_Removed=('vul_count_exceptions_only', lambda x: len(merged_data) - x.sum()),  # Total minus exceptions
    Vul_Count_Closed=('vul_count_closed', 'sum'),
    high_cvss_due_date_0_30=('due date 0-30 days', 'sum'),
    high_cvss_due_date_30_60=('due date 30-60 days', 'sum'),
    high_cvss_due_date_60_90=('due date 60-90 days', 'sum'),
    high_cvss_due_date_gt_90=('due date >90 days', 'sum'),
    high_cvss_exception_exp_0_30=('exception expiration 0-30 days', 'sum'),
    high_cvss_exception_exp_30_60=('exception expiration 30-60 days', 'sum'),
    high_cvss_exception_exp_60_90=('exception expiration 60-90 days', 'sum'),
    high_cvss_exception_exp_gt_90=('exception expiration >90 days', 'sum')
).reset_index()

# Save the final DataFrame to a new Excel file
output_path_final = 'consolidated_vulnerability_data.xlsx'  # Change this to your desired output path
consolidated_data.to_excel(output_path_final, index=False)

print("Consolidated data saved to", output_path_final)
