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


    root = tk.Tk()
    app = GraphGUI(root)
    root.mainloop()


# </ProgramSnippet>


class GraphGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Exchange Verteilerpflege")
        # Create a frame to hold the list of users
        # Create the user list box
        self.user_frame = tk.Frame(self.root)
        self.user_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.user_label = tk.Label(self.user_frame, text="Users")
        self.user_label.pack(side=tk.TOP, padx=5, pady=5)
        self.user_scrollbar = tk.Scrollbar(self.user_frame, orient=tk.VERTICAL)
        self.user_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.user_listbox = tk.Listbox(self.user_frame, yscrollcommand=self.user_scrollbar.set)
        self.user_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.user_scrollbar.config(command=self.user_listbox.yview)

        # Create the group list box
        self.group_frame = tk.Frame(self.root)
        self.group_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.group_label = tk.Label(self.group_frame, text="Groups")
        self.group_label.pack(side=tk.TOP, padx=5, pady=5)
        self.group_scrollbar = tk.Scrollbar(self.group_frame, orient=tk.VERTICAL)
        self.group_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.group_listbox = tk.Listbox(self.group_frame, yscrollcommand=self.group_scrollbar.set)
        self.group_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.group_scrollbar.config(command=self.group_listbox.yview)

        # Add users to the list box
        users = Users()
        users.list_users()
        for user in users.all_users:
            self.user_listbox.insert(tk.END, user['displayName'])

        # Add groups to the list box
        groups = Groups()
        groups.list_groups()
        for group in groups.all_groups:
            self.group_listbox.insert(tk.END, group['displayName'])


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
    all_users = []
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
