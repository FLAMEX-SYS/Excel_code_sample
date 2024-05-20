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

# Calculate aging for high CVSS due dates
for index, row in high_cvss_data.iterrows():
    if pd.notnull(row['Due Date']):
        delta_days = (row['Due Date'] - current_date).days
        if delta_days > 30:
            high_cvss_data.at[index, 'aging due date in 30 plus days high cvss'] = 1
        elif 0 <= delta_days <= 30:
            high_cvss_data.at[index, 'aging due date in 0 -30 days high cvss'] = 1
        elif -30 <= delta_days < 0:
            high_cvss_data.at[index, 'aging past due date 1-30 high cvss'] = 1
        elif -60 <= delta_days < -30:
            high_cvss_data.at[index, 'aging past due date 30 - 60 days high cvss'] = 1
        elif -90 <= delta_days < -60:
            high_cvss_data.at[index, 'aging past due date 60 -90 days high cvss'] = 1
        elif delta_days < -90:
            high_cvss_data.at[index, 'aging past due date 90 to plus high cvss'] = 1

# Calculate aging for exception removed
for index, row in high_cvss_data.iterrows():
    if row['Status'] == 'Removed' and pd.notnull(row['Exception Expiration']):
        delta_days = (row['Exception Expiration'] - current_date).days
        if delta_days > 30:
            high_cvss_data.at[index, 'aging due 30 to plus exceptions removed'] = 1
        elif 0 <= delta_days <= 30:
            high_cvss_data.at[index, 'aging due in 0 to 30 exceptions removed'] = 1
        elif -30 <= delta_days < 0:
            high_cvss_data.at[index, 'aging past due 1-30 exceptions removed'] = 1
        elif -60 <= delta_days < -30:
            high_cvss_data.at[index, 'aging past due 30-60 exceptions removed'] = 1
        elif -90 <= delta_days < -60:
            high_cvss_data.at[index, 'aging past due 60-90 exceptions removed'] = 1
        elif delta_days < -90:
            high_cvss_data.at[index, 'aging past due 90 to plus exceptions removed'] = 1

# Save the modified DataFrame to a new Excel file
output_path_analyzed = 'analyzed_vulnerability_data.xlsx'  # Change this to your desired output path
high_cvss_data.to_excel(output_path_analyzed, index=False)

print("Analyzed data saved to", output_path_analyzed)
