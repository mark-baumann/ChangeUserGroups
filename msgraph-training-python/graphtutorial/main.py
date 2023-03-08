# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# <ProgramSnippet>
import configparser
import graph

def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azureSettings = config['azure']

    initialize_graph(azureSettings)

    greet_user()
    list_users()
    list_groups()
    
# </ProgramSnippet>

# <InitializeGraphSnippet>
def initialize_graph(settings: configparser.SectionProxy):
    graph.initialize_graph_for_user_auth(settings)
# </InitializeGraphSnippet>

# <GreetUserSnippet>
def greet_user():
    user = graph.get_user()
    print('Hello,', user['displayName'])
    # For Work/school accounts, email is in mail property
    # Personal accounts, email is in userPrincipalName
    print('Email:', user['mail'] or user['userPrincipalName'], '\n')
# </GreetUserSnippet>

# <DisplayAccessTokenSnippet>
def display_access_token():
    token = graph.get_user_token()
    print('User token:', token, '\n')
# </DisplayAccessTokenSnippet>

# <ListInboxSnippet>
def list_inbox():
    message_page = graph.get_inbox()

    # Output each message's details
    for message in message_page['value']:
        print('Message:', message['subject'])
        print('  From:', message['from']['emailAddress']['name'])
        print('  Status:', 'Read' if message['isRead'] else 'Unread')
        print('  Received:', message['receivedDateTime'])

    # If @odata.nextLink is present
    more_available = '@odata.nextLink' in message_page
    print('\nMore messages available?', more_available, '\n')
# </ListInboxSnippet>

# <SendMailSnippet>
def send_mail():
    # Send mail to the signed-in user
    # Get the user for their email address
    user = graph.get_user()
    user_email = user['mail'] or user['userPrincipalName']

    graph.send_mail('Testing Microsoft Graph', 'Hello world!', user_email)
    print('Mail sent.\n')
    return
# </SendMailSnippet>

# <ListGroupSnippet>

def list_groups():
    all_groups = graph.get_all_groups()
    
    for group in all_groups:
        print('Group:', group['displayName'])
        print('  ID:', group['id'])
        
    print('\nTotal groups', len(all_groups))

#</ListGroupSnippet>

# <ListUsersSnippet>
def list_users():
    all_users = graph.get_all_users()

    # Output each users's details
    for user in all_users:
        print('User:', user['displayName'])
        print('  ID:', user['id'])
        print('  Email:', user['mail'])

    print('\nTotal users:', len(all_users))
# </ListUsersSnippet>

# <MakeGraphCallSnippet>
def make_graph_call():
    graph.make_graph_call()
# </MakeGraphCallSnippet>

# Run main
main()
