import requests
import pandas as pd

def fetch_all_results(base_url, params, headers):
    all_issues = []
    start_at = 0
    max_results = 100  # Maximum results per request, as defined by the API
    total_results = None

    while total_results is None or start_at < total_results:
        # Update the 'startAt' parameter for pagination
        params['startAt'] = start_at
        
        response = requests.get(base_url, params=params, headers=headers)
        data = response.json()

        # Process and append the current batch of issues
        issues = data['issues']
        all_issues.extend(issues)

        # Update counters and conditions for the loop
        if total_results is None:
            total_results = data['total']
        start_at += len(issues)

    return all_issues

# Base URL for the API endpoint
base_url = "https://your-domain.atlassian.net/rest/api/3/search"

# Initial parameters and headers (e.g., authentication)
params = {
    "jql": "project = YOURPROJECTKEY",
    "maxResults": 100,
    "fields": "id,key,fields"
}
headers = {
    "Authorization": "Basic YOUR_ENCODED_CREDENTIALS",
    "Content-Type": "application/json"
}

# Fetch all results
all_issues = fetch_all_results(base_url, params, headers)

# Convert to DataFrame and process as shown in previous examples
# Assuming 'all_issues' is now a list of all issues fetched through pagination
df_issues = pd.json_normalize(all_issues)
# Continue with processing and exporting to CSV...
