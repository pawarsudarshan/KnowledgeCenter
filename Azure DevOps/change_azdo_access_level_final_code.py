import requests
import base64
import json
import pandas as pd

# Constants
organization = 'Byaghru'
pat = 'YOUR_PAT_HERE'
excel_file_path = '/Byaghru_users.xlsx'
count = 0
errorCount = 0

# Encode PAT for Basic Auth
pat_encoded = base64.b64encode(bytes(f':{pat}', 'utf-8')).decode('utf-8')

# Azure DevOps API URL to get user descriptor
user_descriptor_url = f'https://vsaex.dev.azure.com/{organization}/_apis/userentitlements?api-version=6.0-preview.3'

# Headers for getting user descriptor
headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Basic {pat_encoded}'
}

# Headers for the PATCH request
patch_headers = {
    'Content-Type': 'application/json-patch+json',
    'Authorization': f'Basic {pat_encoded}'
}

# Read email IDs from Excel file
try:
    emails_df = pd.read_excel(excel_file_path)
except Exception as e:
    print(f'Error reading the Excel file: {e}')
    exit()

email_column_name = 'Email'  # Change this to the name of the column containing the email IDs
current_access_column_name = 'Current Access Level'
change_access_to_column_name = 'Change Access To'

# Get all user descriptors once
response = requests.get(user_descriptor_url, headers=headers)
if response.status_code != 200:
    print(f'Failed to get user descriptors. Status code: {response.status_code}')
    print(response.text)
    exit()

users = response.json().get('members', [])
user_descriptors = {user['user']['principalName'].lower(): user['id'] for user in users}

# Define a mapping of access levels to their corresponding API values
access_level_mapping = {
    'Stakeholder': 'stakeholder',
    'Basic': 'express',
    'VisualStudioSubscription': 'advanced'
}

for index, row in emails_df.iterrows():
    user_email = row[email_column_name].strip().lower()
    user_current_access_level = row[current_access_column_name]
    user_change_to_access_level = row[change_access_to_column_name]

    user_descriptor = user_descriptors.get(user_email)
    if not user_descriptor:
        print(f'User with email {user_email} not found.')
        continue

    # Azure DevOps API URL to update user access level
    update_url = f'https://vsaex.dev.azure.com/{organization}/_apis/userentitlements/{user_descriptor}?api-version=6.0-preview.3'

    # Get the correct access level value for the API
    new_access_level = access_level_mapping.get(user_change_to_access_level)
    if not new_access_level:
        errorCount += 1
        print(f'{errorCount}: Invalid access level: {user_change_to_access_level} for user {user_email}')
        continue

    # Payload with new access level
    payload = [
        {
            "op": "replace",
            "path": "/accessLevel",
            "value": {
                "accountLicenseType": new_access_level
            }
        }
    ]

    # Update the user's access level
    response = requests.patch(update_url, headers=patch_headers, data=json.dumps(payload))
    # print(response.text)

    if response.status_code == 200:
        count += 1
        print(f'{count}. Successfully updated access level for {user_email} from {user_current_access_level} to {user_change_to_access_level}')
    else:
        print(f'Failed to update access level for {user_email}. Status code: {response.status_code}')
        print(response.text)

print(f"Total {count} users successfully changed")
print(f"Total {errorCount} users failed")
