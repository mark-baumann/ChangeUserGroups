# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# <UserAuthConfigSnippet>
import sys
import json
from azure.identity import DeviceCodeCredential, ClientSecretCredential
from msgraph.core import GraphClient

# Assign variables to the module so they stay set
this = sys.modules[__name__]

def initialize_graph_for_user_auth(config):
    this.settings = config
    client_id = this.settings['clientId']
    tenant_id = this.settings['authTenant']
    graph_scopes = this.settings['graphUserScopes'].split(' ')

    this.device_code_credential = DeviceCodeCredential(client_id, tenant_id = tenant_id)
    this.user_client = GraphClient(credential=this.device_code_credential, scopes=graph_scopes)
# </UserAuthConfigSnippet>

# <GetUserTokenSnippet>
def get_user_token():
    graph_scopes = this.settings['graphUserScopes']
    access_token = this.device_code_credential.get_token(graph_scopes)
    return access_token.token
# </GetUserTokenSnippet>

# <GetUserSnippet>
def get_user():
    endpoint = '/me'
    # Only request specific properties
    select = 'displayName,mail,userPrincipalName'
    request_url = f'{endpoint}?$select={select}'

    user_response = this.user_client.get(request_url)
    return user_response.json()
# </GetUserSnippet>

# <GetInboxSnippet>
def get_inbox():
    endpoint = '/me/mailFolders/inbox/messages'
    # Only request specific properties
    select = 'from,isRead,receivedDateTime,subject'
    # Get at most 25 results
    top = 25
    # Sort by received time, newest first
    order_by = 'receivedDateTime DESC'
    request_url = f'{endpoint}?$select={select}&$top={top}&$orderBy={order_by}'

    inbox_response = this.user_client.get(request_url)
    return inbox_response.json()
# </GetInboxSnippet>

# <SendMailSnippet>
def send_mail(subject: str, body: str, recipient: str):
    request_body = {
        'message': {
            'subject': subject,
            'body': {
                'contentType': 'text',
                'content': body
            },
            'toRecipients': [
                {
                    'emailAddress': {
                        'address': recipient
                    }
                }
            ]
        }
    }

    request_url = '/me/sendmail'

    this.user_client.post(request_url,
                          data=json.dumps(request_body),
                          headers={'Content-Type': 'application/json'})
# </SendMailSnippet>

# <AppOnyAuthConfigSnippet>
def ensure_graph_for_app_only_auth():
    if not hasattr(this, 'client_credential'):
        client_id = this.settings['clientId']
        tenant_id = this.settings['tenantId']
        client_secret = this.settings['clientSecret']

        this.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)

    if not hasattr(this, 'app_client'):
        this.app_client = GraphClient(credential=this.client_credential,
                                      scopes=['https://graph.microsoft.com/.default'])
# </AppOnyAuthConfigSnippet>

# <GetUsersSnippet>
def get_all_users():
    ensure_graph_for_app_only_auth()

    endpoint = '/users'
    # Only request specific properties
    select = 'displayName,id,mail'
    # Get at most 25 results per page
    top = 25
    # Sort by display name
    order_by = 'displayName'
    # Initialize empty list to store all users
    all_users = []

    # Make initial request for first page of results
    request_url = f'{endpoint}?$select={select}&$top={top}&$orderBy={order_by}'
    response = this.app_client.get(request_url)
    response_data = response.json()

    # Keep requesting next page until there are no more results
    while '@odata.nextLink' in response_data:
        all_users.extend(response_data['value'])
        next_link = response_data['@odata.nextLink']
        response = this.app_client.get(next_link)
        response_data = response.json()

    # Add final page of results to the list
    all_users.extend(response_data['value'])

    return all_users
# </GetUsersSnippet>



# <GetGroupCallSnippet>
def get_all_groups():
    ensure_graph_for_app_only_auth()
    
    endpoint = '/groups'
    # Only request specific properties
    select = 'displayName,id,description,groupTypes'
    # Get at most 25 results per page
    top = 25
    # Sort by display name
    order_by = 'displayName'
    # Initialize empty list to store all groups
    all_groups = []

    # Make initial request for first page of results
    request_url = f'{endpoint}?$select={select}&$top={top}&$orderBy={order_by}'
    response = this.app_client.get(request_url)
    response_data = response.json()

    # Keep requesting next page until there are no more results
    while '@odata.nextLink' in response_data:
        all_groups.extend(response_data['value'])
        next_link = response_data['@odata.nextLink']
        response = this.app_client.get(next_link)
        response_data = response.json()

    # Add final page of results to the list
    all_groups.extend(response_data['value'])

    return all_groups

# </GetGroupCallSnippet>

# <MakeGraphCallSnippet>
def make_graph_call():
    # INSERT YOUR CODE HERE
    # Note: if using app_client, be sure to call
    # ensure_graph_for_app_only_auth before using it
    # ensure_graph_for_app_only_auth()
    return
# </MakeGraphCallSnippet>


def add_member_to_group(user_id, group_id):

    ensure_graph_for_app_only_auth()

    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members/$ref"
    headers = {"Content-Type": "application/json"}
    data = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"}
    response = this.app_client.post(url, headers=headers, json=data)
    if response.status_code == 204:
        print(f'Successfully added user {user_id} to group {group_id}.')
    else:
        print('Error:', response.content)