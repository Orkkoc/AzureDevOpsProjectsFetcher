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
    url = f'https://{organization}.visualstudio.com/{collection}/_apis/projects'
    response = requests.get(url, headers=headers, auth=credentials)
    projects = response.json()['value']

    # Check for continuation token and fetch more projects if available
    while 'x-ms-continuationtoken' in response.headers:
        continuation_token = response.headers['x-ms-continuationtoken']
        next_url = f'{url}?continuationToken={continuation_token}'
        response = requests.get(next_url, headers=headers, auth=credentials)
        projects.extend(response.json()['value'])

    return projects

def get_users_for_project(collection, project_id):
    """
    Fetches users associated with a given project.
    """
    users = []
    url = f'https://{organization}.visualstudio.com/{collection}/_apis/projects/{project_id}/teams?api-version=5.1'
    response = requests.get(url, headers=headers, auth=credentials)
    
    if response.status_code == 200:
        teams = response.json()['value']
        for team in teams:
            team_id = team['id']
            team_users_url = f'https://{organization}.visualstudio.com/{collection}/_apis/projects/{project_id}/teams/{team_id}/members?api-version=5.1'
            team_users_response = requests.get(team_users_url, headers=headers, auth=credentials)
            
            if team_users_response.status_code == 200:
                team_members = team_users_response.json()['value']
                for member in team_members:
                    # Extract LDAP name and other details
                    user_details = {
                        'LDAPName': member['identity']['uniqueName'],
                        'displayName': member['identity']['displayName'],
                        'userId': member['identity']['id'],
                        'projectId': project_id,
                        'projectName': project['name'],
                        'collectionName': collection  # Add collection name
                    }
                    users.append(user_details)

    return users

all_projects = []
all_users = []
for collection in collections:
    projects = get_projects(collection)
    for project in projects:
        project_id = project['id']
        users = get_users_for_project(collection, project_id)  # Pass collection name
        for user in users:
            user['projectId'] = project_id
            user['projectName'] = project['name']
            all_users.append(user)
        project['Collection'] = collection
    all_projects.extend(projects)  # Move this outside the inner loop

from datetime import datetime
# Current timestamp
timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')

# Filenames with timestamp
projects_filename = f'azure_devops_projects_{timestamp}.xlsx'
users_filename = f'azure_devops_project_users_{timestamp}.xlsx'

# Convert to DataFrame and save to Excel with timestamped filenames
df_projects = pd.DataFrame(all_projects)
df_projects.to_excel(projects_filename, index=False)

df_users = pd.DataFrame(all_users)
df_users.to_excel(users_filename, index=False)
