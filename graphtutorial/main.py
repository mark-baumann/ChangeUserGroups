# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# <ProgramSnippet>
import configparser
import graph
import tkinter as tk


def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azureSettings = config['azure']
    
    initialize_graph(azureSettings)
    
    greet_user()
    
    users = Users()
    users.list_users()
    
    groups = Groups()
    groups.list_groups()
    
    window = tk.Tk()
    window.title("My Graph GUI")
    window.mainloop()
# </ProgramSnippet>


class Groups:
# <ListGroupSnippet>
    all_groups = []
    def list_groups(self):
        self.all_groups = graph.get_all_groups()
        
        for group in self.all_groups:
            print('Group:', group['displayName'])
            print('  ID:', group['id'])
            
        print('\nTotal groups', len(self.all_groups))
    #</ListGroupSnippet>

class Users:
    all_groups = []
    # <ListUsersSnippet>
    def list_users(self):
        self.all_users = graph.get_all_users()

        # Output each users's details
        for user in self.all_users:
            print('User:', user['displayName'])
            print('  ID:', user['id'])
            print('  Email:', user['mail'])

        print('\nTotal users:', len(self.all_users))
    # </ListUsersSnippet>


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


# Run main
main()
