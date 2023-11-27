# Azure DevOps Projects Fetcher

This script fetches projects from multiple Azure DevOps collections and saves them into an Excel file.

## Installation

To run this script, you need Python installed on your system. If you don't have Python, you can download it from [python.org](https://www.python.org/downloads/).

### Required Python Packages

- `requests`: For making API requests.
- `pandas`: For data manipulation and exporting to Excel.
- `openpyxl`: For handling Excel file operations in Pandas.

You can install these packages using pip:

```bash
pip install requests pandas openpyxl
```
## Usage

To use this script, follow these steps:

1. Ensure Python and the necessary packages (`requests`, `pandas`, and `openpyxl`) are installed on your system.
2. Clone this repository or download the `get_azure_projects.py` script to your local machine.
3. Open the script in a text editor and update the `organization` and `personal_access_token` variables with your Azure DevOps organization name and Personal Access Token.
4. Modify the `collections` list in the script to include the Azure DevOps collections from which you want to fetch projects.
5. Run the script using a Python interpreter. You can do this from a command line or terminal:

```bash
python get_azure_projects.py
```
6. The script will fetch projects from the specified collections and save them into an Excel file named azure_devops_projects.xlsx.

# Azure DevOps Projects and Users Fetcher

This script fetches projects and their associated users from multiple Azure DevOps collections and saves them into timestamped Excel files.

## Installation

To run this script, you need Python installed on your system. If you don't have Python, you can download it from [python.org](https://www.python.org/downloads/).

### Required Python Packages

- `requests`: For making API requests.
- `pandas`: For data manipulation and exporting to Excel.
- `openpyxl`: For handling Excel file operations in Pandas.

You can install these packages using pip:

```bash
pip install requests pandas openpyxl
```
Usage
To use this script, follow these steps:

1. Ensure Python and the necessary packages (requests, pandas, and openpyxl) are installed on your system.
2. Clone this repository or download the get_azure_projects_users.py script to your local machine.
3. Open the script in a text editor and update the organization and personal_access_token variables with your Azure DevOps organization name and Personal Access Token.
4. Modify the collections list in the script to include the Azure DevOps collections from which you want to fetch projects and users.
5. Run the script using a Python interpreter. You can do this from a command line or terminal:

```bash
python get_azure_projects_users.py
```
6. The script will fetch projects and their users from the specified collections. It will create two timestamped Excel files, one for projects and one for users, named in the format azure_devops_projects_YYYYMMDD-HHMMSS.xlsx and azure_devops_project_users_YYYYMMDD-HHMMSS.xlsx.

#### License

```markdown
## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```
