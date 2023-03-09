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
        self.root.geometry("1000x1000")

        # Create a frame to hold the user and group dropdowns
        self.dropdown_frame = tk.Frame(self.root)
        self.dropdown_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Create the user dropdown
        self.user_label = tk.Label(self.dropdown_frame, text="User")
        self.user_label.pack(side=tk.LEFT, padx=5, pady=5)
        self.user_var = tk.StringVar(self.dropdown_frame)
        self.user_dropdown = tk.OptionMenu(self.dropdown_frame, self.user_var, '')
        self.user_dropdown.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create the group dropdown
        self.group_label = tk.Label(self.dropdown_frame, text="Group")
        self.group_label.pack(side=tk.LEFT, padx=5, pady=5)
        self.group_var = tk.StringVar(self.dropdown_frame)
        self.group_dropdown = tk.OptionMenu(self.dropdown_frame, self.group_var, '')
        self.group_dropdown.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add users to the user dropdown
        self.users = Users()
        self.users.list_users()
        user_options = []
        for user in self.users.all_users:
            user_options.append(user['displayName'])
        self.user_dropdown['menu'].delete(0, tk.END)
        for option in user_options:
            self.user_dropdown['menu'].add_command(label=option, command=tk._setit(self.user_var, option))

        # Add groups to the group dropdown
        self.groups = Groups()
        self.groups.list_groups()
        group_options = []
        for group in self.groups.all_groups:
            group_options.append(group['displayName'])
        self.group_dropdown['menu'].delete(0, tk.END)
        for option in group_options:
            self.group_dropdown['menu'].add_command(label=option, command=tk._setit(self.group_var, option))

        # Create a frame to hold the Add and Delete buttons
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Create the Add button
        self.add_button = tk.Button(self.button_frame, text="Add", command=self.add_user_to_group)
        self.add_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Create the Delete button
        self.delete_button = tk.Button(self.button_frame, text="Delete", command=self.remove_user_from_group)
        self.delete_button.pack(side=tk.LEFT, padx=5, pady=5)

    def add_user_to_group(self):
        user_name = self.user_var.get()
        group_name = self.group_var.get()
        user_id = ''
        group_id = ''
        for user in self.users.all_users:
            if user['displayName'] == user_name:
                user_id = user['id']
                break
        for group in self.groups.all_groups:
            if group['displayName'] == group_name:
                print(group['displayName'])
                group_id = group['id']
                break
        if user_id and group_id:
            graph.add_member_to_group(user_id, group_id)
            print(f"Added user {user_name} ({user_id}) to group {group_name} ({group_id})")
        else:
            print("Could not find user or group")
            
    def remove_user_from_group(self):
        pass

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