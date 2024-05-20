import pandas as pd
from datetime import datetime

# Load the vulnerability data
file_path = 'path_to_your_vulnerability_data.xlsx'
vul_data = pd.read_excel(file_path)

# Convert columns to datetime, ensuring correct date parsing
vul_data['Due Date'] = pd.to_datetime(vul_data['Due Date'])
vul_data['Exception Expiration'] = pd.to_datetime(vul_data['Exception Expiration'])

# Define today's date for aging calculations
today = pd.Timestamp('today')

# Define the high CVSS threshold
high_cvss_threshold = 7

# Define the columns for the summary DataFrame explicitly
summary_columns = [
    'hostname', 'vul count', 'vul count high cvss', 'vul count exception removed', 
    'vul count exception only', 'vul count closed', 'aging due in 30 plus days high cvss', 
    'aging due in 0 -30 days high cvss', 'aging past due 1-30 high cvss', 
    'aging past due 30 - 60 days high cvss', 'aging past due 60 -90 days high cvss', 
    'aging past due 90 to plus high cvss', 'aging due 30 to plus exceptions removed', 
    'aging due in 0 to 30 exceptions removed', 'aging past due 1-30 exceptions removed', 
    'aging past due 30-60 exceptions removed', 'aging past due 60-90 exceptions removed', 
    'aging past due 90 to plus exceptions removed'
]
summary_df = pd.DataFrame(columns=summary_columns)

# Group by hostname
grouped = vul_data.groupby('hostname')

for name, group in grouped:
    data = {col: 0 for col in summary_columns}  # Initialize all to zero
    data['hostname'] = name
    data['vul count'] = len(group)
    data['vul count high cvss'] = len(group[group['CVSS'] > high_cvss_threshold])
    data['vul count exception removed'] = len(group[group['Status'] == 'Removed'])
    data['vul count exception only'] = len(group[group['Exception Number'].notna() & (group['Status'] != 'Removed')])
    data['vul count closed'] = len(group[group['Status'] == 'Closed'])

    # High CVSS aging calculations
    for row in group.itertuples():
        if row.CVSS > high_cvss_threshold:
            delta_days = (today - row.Due_Date).days
            if delta_days > 90:
                data['aging past due 90 to plus high cvss'] += 1
            elif delta_days > 60:
                data['aging past due 60-90 days high cvss'] += 1
            elif delta_days > 30:
                data['aging past due 30 - 60 days high cvss'] += 1
            elif delta_days > 0:
                data['aging past due 1-30 high cvss'] += 1
            elif delta_days > -30:
                data['aging due in 0 -30 days high cvss'] += 1
            else:
                data['aging due in 30 plus days high cvss'] += 1

    # Exception removed aging calculations
    for row in group[group['Status'] == 'Removed'].itertuples():
        delta_days = (today - row.Exception_Expiration).days
        if delta_days > 90:
            data['aging past due 90 to plus exceptions removed'] += 1
        elif delta_days > 60:
            data['aging past due 60-90 exceptions removed'] += 1
        elif delta_days > 30:
            data['aging past due 30-60 exceptions removed'] += 1
        elif delta_days > 0:
            data['aging past due 1-30 exceptions removed'] += 1
        elif delta_days > -30:
            data['aging due in 0 to 30 exceptions removed'] += 1
        else:
            data['aging due 30 to plus exceptions removed'] += 1

    # Append data for this hostname to the summary DataFrame
    summary_df = summary_df.append(data, ignore_index=True)

# Save the summary DataFrame to a new Excel file
output_path_summary = 'vulnerability_summary.xlsx'
summary_df.to_excel(output_path_summary, index=False)

print("Vulnerability summary saved to", output_path_summary)
