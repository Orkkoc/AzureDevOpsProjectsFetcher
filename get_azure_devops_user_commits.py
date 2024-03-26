import requests
import pandas as pd
from datetime import datetime

# Azure DevOps details
personal_access_token = 'YOUR_PERSONAL_ACCESS_TOKEN'
organization_url = 'YOUR_ORGANIZATION_URL'
author_name = 'AUTHOR_NAME'
collections = [
    'Collection1', 'Collection2', 'Collection3'
]

# Setup for Basic Authentication
credentials = ('', personal_access_token)
headers = {'Content-Type': 'application/json'}

def get_repositories(collection, project_id):
    """
    Fetches repositories associated with a given project.
    """
    repos_url = f'{organization_url}/{collection}/{project_id}/_apis/git/repositories?api-version=5.1'
    response = requests.get(repos_url, headers=headers, auth=credentials)
    if response.status_code == 200 and 'value' in response.json():
        return response.json()['value']
    else:
        print(f"Failed to retrieve repositories for project {project_id} in collection {collection}.")
        return []

def get_commits_for_user(collection, project_name, repo_name, repo_id, author_name):
    """
    Fetches commits made by a specific user in a given repository.
    """
    commits_url = f'{organization_url}/{collection}/{project_id}/_apis/git/repositories/{repo_id}/commits?api-version=5.1&searchCriteria.author={author_name}'
    response = requests.get(commits_url, headers=headers, auth=credentials)
    commits = []
    if response.status_code == 200 and 'value' in response.json():
        for commit in response.json()['value']:
            commit['project_name'] = project_name  # Add project name
            commit['repository_name'] = repo_name  # Add repository name
            commits.append(commit)
    else:
        print(f"Failed to retrieve commits for repository {repo_name} in project {project_name}.")
    return commits

def get_tfvc_changesets(collection, project_name, project_id, author_name):
    """
    Fetches TFVC changesets for a given project, filtered by author.
    """
    changesets_url = f'{organization_url}/{collection}/_apis/tfvc/changesets?searchCriteria.itemPath=$/{project_id}&searchCriteria.author={author_name}&api-version=5.1'
    response = requests.get(changesets_url, headers=headers, auth=credentials)
    changesets = []
    if response.status_code == 200 and 'value' in response.json():
        for changeset in response.json()['value']:
            changeset['project_name'] = project_name  # Add project name
            changesets.append(changeset)
    else:
        print(f"Failed to fetch TFVC changesets for project {project_name}.")
    return changesets

all_data = []
for collection in collections:
    projects_url = f'{organization_url}/{collection}/_apis/projects?api-version=5.1'
    response = requests.get(projects_url, headers=headers, auth=credentials)

    if response.status_code == 200 and 'value' in response.json():
        projects = response.json()['value']
        for project in projects:
            project_id = project['id']
            project_name = project['name']

            # Fetch Git repositories and their commits
            repos = get_repositories(collection, project_id)
            for repo in repos:
                repo_id = repo['id']
                repo_name = repo['name']
                commits = get_commits_for_user(collection, project_name, repo_name, repo_id, author_name)
                if commits:
                    all_data.extend(commits)

            # Fetch TFVC changesets
            tfvc_changesets = get_tfvc_changesets(collection, project_name, project_id, author_name)
            if tfvc_changesets:
                all_data.extend(tfvc_changesets)
    else:
        print(f"Failed to retrieve projects for collection {collection}. Response: {response.text}")

# Generate datetime stamp for the file name
datetime_stamp = datetime.now().strftime("%Y%m%d%H%M%S")
file_name = f'commits_from_{author_name}_{datetime_stamp}.xlsx'

# Convert to DataFrame and save to an Excel file
df_commits = pd.DataFrame(all_data)
df_commits.to_excel(file_name, index=False)

print(f"Data exported to {file_name}")
