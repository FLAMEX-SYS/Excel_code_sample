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

# Filter for high CVSS scores (greater than 7)
high_cvss_data = extracted_data[extracted_data['CVSS'] > 7]

# Reset index to ensure unique keys
high_cvss_data.reset_index(drop=True, inplace=True)

# Initialize new columns with blank values
aging_columns = [
    'aging due date in 30 plus days high cvss', 
    'aging due date in 0 -30 days high cvss',
    'aging past due date 1-30 high cvss', 
    'aging past due date 30 - 60 days high cvss', 
    'aging past due date 60 -90 days high cvss', 
    'aging past due date 90 to plus high cvss',
    'aging due 30 to plus exceptions removed', 
    'aging due in 0 to 30 exceptions removed', 
    'aging past due 1-30 exceptions removed', 
    'aging past due 30-60 exceptions removed', 
    'aging past due 60-90 exceptions removed', 
    'aging past due 90 to plus exceptions removed'
]

for column in aging_columns:
    high_cvss_data[column] = ""

# Function to categorize aging
def categorize_aging(row, date_column, category_prefix):
    if pd.notnull(row[date_column]):
        delta_days = (row[date_column] - current_date).days
        if delta_days > 30:
            return category_prefix + ' due date in 30 plus days'
        elif 0 <= delta_days <= 30:
            return category_prefix + ' due date in 0 -30 days'
        elif -30 <= delta_days < 0:
            return category_prefix + ' past due date 1-30'
        elif -60 <= delta_days < -30:
            return category_prefix + ' past due date 30 - 60 days'
        elif -90 <= delta_days < -60:
            return category_prefix + ' past due date 60 -90 days'
        elif delta_days < -90:
            return category_prefix + ' past due date 90 to plus'
    return ""

# Calculate aging for high CVSS due dates
for index, row in high_cvss_data.iterrows():
    high_cvss_category = categorize_aging(row, 'Due Date', 'aging')
    if high_cvss_category:
        high_cvss_data.at[index, high_cvss_category] = 1

# Calculate aging for exception removed
for index, row in high_cvss_data.iterrows():
    if row['Status'] == 'Removed':
        exception_category = categorize_aging(row, 'Exception Expiration', 'aging due')
        if exception_category:
            high_cvss_data.at[index, exception_category + ' exceptions removed'] = 1

# Save the modified DataFrame to a new Excel file
output_path_analyzed = 'analyzed_vulnerability_data.xlsx'  # Change this to your desired output path
high_cvss_data.to_excel(output_path_analyzed, index=False)

print("Analyzed data saved to", output_path_analyzed)
