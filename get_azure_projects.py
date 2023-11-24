import requests
import pandas as pd

# Azure DevOps details
# Replace these with your organization name and Personal Access Token
organization = 'your_organization'
personal_access_token = 'your_pat'

# List of collections to fetch projects from
collections = ['DefaultCollection', 'SecondCollection']

# Setup for Basic Authentication
credentials = ('', personal_access_token)
headers = {'Content-Type': 'application/json'}

def get_projects(collection):
    """
    Fetches projects from a given Azure DevOps collection.
    Handles pagination to retrieve all projects.
    """
    url = f'https://{organization}.visualstudio.com/{collection}/_apis/projects'
    response = requests.get(url, headers=headers, auth=credentials)
    projects = response.json()['value']

    # Handle pagination
    while 'x-ms-continuationtoken' in response.headers:
        continuation_token = response.headers['x-ms-continuationtoken']
        next_url = f'{url}?continuationToken={continuation_token}'
        response = requests.get(next_url, headers=headers, auth=credentials)
        projects.extend(response.json()['value'])

    return projects

all_projects = []
for collection in collections:
    # Fetch projects for each collection and add them to the all_projects list
    projects = get_projects(collection)
    for project in projects:
        project['Collection'] = collection  # Add collection name for identification
    all_projects.extend(projects)

from datetime import datetime
# Current timestamp
timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')

# Filenames with timestamp
projects_filename = f'azure_devops_projects_{timestamp}.xlsx'

# Convert to DataFrame and save to Excel with timestamped filenames
df_projects = pd.DataFrame(all_projects)
df_projects.to_excel(projects_filename, index=False)
