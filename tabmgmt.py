import os
import configparser
import threading
import tkinter as tk
from tkinter import ttk, messagebox, Listbox, Scrollbar
import pandas as pd
from datetime import datetime, timezone
import tableauserverclient as tsc
from ldap3 import Server, Connection, ALL, SUBTREE

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tableau Site Management App")
        self.geometry("600x600")
        self.config(bg='white')
        self.config = self.load_config()
        self.auth = None
        self.server = None
        if not self.config_is_valid():
            self.show_config_frame(from_main=False)
        else:
            self.load_auth()
            self.show_main_frame()

    def load_config(self):
        path = os.path.join(os.environ['USERPROFILE'], 'tabmgt.env')
        config = configparser.ConfigParser()
        config.read(path)
        if 'DEFAULT' not in config:
            config['DEFAULT'] = {}
        return config

    def save_config(self):
        path = os.path.join(os.environ['USERPROFILE'], 'tabmgt.env')
        with open(path, 'w') as f:
            self.config.write(f)

    def config_is_valid(self):
        req = ['Tableau Server URL', 'Tableau Site Name', 'Access Token Name', 'Access Token', 'Email Address', 'NTLogin',
               'LDAP Server', 'LDAP Port', 'LDAP User', 'LDAP Password', 'LDAP Base DN']
        for r in req:
            if r not in self.config['DEFAULT'] or not self.config['DEFAULT'][r]:
                return False
        return True

    def load_auth(self):
        d = self.config['DEFAULT']
        self.server_url = d['Tableau Server URL']
        self.site_id = d['Tableau Site Name']
        self.token_name = d['Access Token Name']
        self.token_value = d['Access Token']
        self.email = d['Email Address']
        self.ntlogin = d['NTLogin']
        self.ldap_server = d['LDAP Server']
        self.ldap_port = int(d['LDAP Port'])
        self.ldap_user = d['LDAP User']
        self.ldap_password = d['LDAP Password']
        self.ldap_base_dn = d['LDAP Base DN']
        self.auth = tsc.PersonalAccessTokenAuth(self.token_name, self.token_value, self.site_id)
        self.server = tsc.Server(self.server_url)
        # self.server.add_http_options({'verify': False})  # Uncomment if certificate issues

    def test_auth(self):
        try:
            with self.server.auth.sign_in(self.auth):
                return True
        except:
            return False

    def show_config_frame(self, from_main):
        if hasattr(self, 'current_frame'):
            self.current_frame.destroy()
        self.current_frame = ConfigFrame(self, from_main)
        self.current_frame.pack(fill='both', expand=True)

    def show_main_frame(self):
        if hasattr(self, 'current_frame'):
            self.current_frame.destroy()
        self.current_frame = MainFrame(self)
        self.current_frame.pack(fill='both', expand=True)

    def show_frame(self, frame_class):
        if hasattr(self, 'current_frame'):
            self.current_frame.destroy()
        self.current_frame = frame_class(self)
        self.current_frame.pack(fill='both', expand=True)

class ConfigFrame(tk.Frame):
    def __init__(self, parent, from_main):
        super().__init__(parent, bg='white')
        self.parent = parent
        self.from_main = from_main
        vars = ['Tableau Server URL', 'Tableau Site Name', 'Access Token Name', 'Access Token', 'Email Address', 'NTLogin',
                'LDAP Server', 'LDAP Port', 'LDAP User', 'LDAP Password', 'LDAP Base DN']
        self.entries = {}
        for i, v in enumerate(vars):
            tk.Label(self, text=v, bg='white').grid(row=i, column=0, padx=10, pady=5)
            e = tk.Entry(self)
            e.grid(row=i, column=1, padx=10, pady=5)
            if v in parent.config['DEFAULT']:
                e.insert(0, parent.config['DEFAULT'][v])
            self.entries[v] = e
        btn_frame = tk.Frame(self, bg='white')
        btn_frame.grid(row=len(vars), column=0, columnspan=2)
        if from_main:
            tk.Button(btn_frame, text="Back", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=self.back).pack(side='left', padx=10)
        tk.Button(btn_frame, text="Submit", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=self.submit).pack(side='left', padx=10)

    def back(self):
        self.parent.show_main_frame()

    def submit(self):
        for k, e in self.entries.items():
            self.parent.config['DEFAULT'][k] = e.get()
        self.parent.save_config()
        self.parent.load_auth()
        if self.parent.test_auth():
            messagebox.showinfo("Success", "Configuration saved and authenticated.")
            self.parent.show_main_frame()
        else:
            messagebox.showerror("Error", "Authentication failed.")

class MainFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        tk.Button(self, text="⚙", command=self.config, bg='gray', fg='white', font=('Arial', 12, 'bold')).pack(anchor='ne')
        buttons = [
            ("User Management", self.user_mgt, '#005F9E'),
            ("Workbook Management", self.placeholder, 'gray'),
            ("Schedule Management", self.placeholder, 'gray'),
            ("Connection Management", self.connection_mgt, '#005F9E'),
            ("Generate Reports", self.reports, '#005F9E'),
        ]
        for text, cmd, color in buttons:
            tk.Button(self, text=text, bg=color, fg='white', font=('Arial', 12, 'bold'), command=cmd).pack(fill='x', pady=10, padx=20)

    def config(self):
        self.parent.show_config_frame(from_main=True)

    def user_mgt(self):
        self.parent.show_frame(UserManagementFrame)

    def connection_mgt(self):
        self.parent.show_frame(ConnectionManagementFrame)

    def placeholder(self):
        messagebox.showinfo("Placeholder", "Not implemented yet.")

    def reports(self):
        self.parent.show_frame(ReportsFrame)

class ConnectionManagementFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        buttons = [
            ("Update Login Details", self.placeholder, 'gray'),
            ("Update Server Name", self.update_server_name, '#005F9E'),
            ("Back", self.back, '#005F9E'),
        ]
        for text, cmd, color in buttons:
            tk.Button(self, text=text, bg=color, fg='white', font=('Arial', 12, 'bold'), command=cmd).pack(fill='x', pady=10, padx=20)

    def back(self):
        self.parent.show_main_frame()

    def update_server_name(self):
        self.parent.show_frame(UpdateServerNameFrame)

    def placeholder(self):
        messagebox.showinfo("Placeholder", "Not implemented yet.")

class UpdateServerNameFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        self.projects = []
        try:
            with parent.server.auth.sign_in(parent.auth):
                parent.server.use_highest_version()
                self.projects = sorted([p.name for p in tsc.Pager(parent.server.projects)])
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.back()
            return

        tk.Label(self, text="Old Server", bg='white').pack(pady=5)
        self.old_server_entry = tk.Entry(self)
        self.old_server_entry.pack(pady=5)

        tk.Label(self, text="New Server", bg='white').pack(pady=5)
        self.new_server_entry = tk.Entry(self)
        self.new_server_entry.pack(pady=5)

        tk.Label(self, text="Select Projects", bg='white').pack(pady=5)
        scrollbar = Scrollbar(self)
        self.project_list = Listbox(self, selectmode='multiple', yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.project_list.yview)
        for p in self.projects:
            self.project_list.insert('end', p)
        self.project_list.pack(side='left', fill='both', expand=True, pady=5, padx=10)
        scrollbar.pack(side='right', fill='y')

        tk.Button(self, text="Submit", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=self.submit).pack(pady=5)
        tk.Button(self, text="Back", bg='#005F9E', fg='white', font=('Arial', 12, 'bold'), command=self.back).pack(pady=5)

    def back(self):
        self.parent.show_frame(ConnectionManagementFrame)

    def submit(self):
        old_server = self.old_server_entry.get().strip()
        new_server = self.new_server_entry.get().strip()
        selected_indices = self.project_list.curselection()
        selected_projects = [self.projects[i] for i in selected_indices]

        if not selected_projects:
            messagebox.showerror("Error", "At least one project must be selected.")
            return
        if not old_server:
            messagebox.showerror("Error", "Old Server name cannot be empty.")
            return
        if not new_server:
            messagebox.showerror("Error", "New Server name cannot be empty.")
            return

        try:
            with self.parent.server.auth.sign_in(self.parent.auth):
                self.parent.server.use_highest_version()
                projects = [p for p in tsc.Pager(self.parent.server.projects) if p.name in selected_projects]
                project_ids = [p.id for p in projects]
                workbooks_updated = 0

                for workbook in tsc.Pager(self.parent.server.workbooks):
                    if workbook.project_id not in project_ids:
                        continue
                    self.parent.server.workbooks.populate_connections(workbook)
                    updated = False
                    for connection in workbook.connections:
                        if connection.datasource_id:
                            continue
                        if old_server == '*' or (connection.server_address and old_server.lower() in connection.server_address.lower()):
                            connection.server_address = new_server
                            updated = True
                    if updated:
                        self.parent.server.workbooks.update_connection(workbook, workbook.connections)
                        workbooks_updated += 1

            messagebox.showinfo("Success", f"Updated {workbooks_updated} workbooks successfully.")
            self.back()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update connections: {str(e)}")

class UserManagementFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        buttons = [
            ("Add User", self.add_user, '#005F9E'),
            ("Remove User", self.remove_user, '#005F9E'),
            ("Manage User Groups", self.placeholder, 'gray'),
            ("Back", self.back, '#005F9E'),
        ]
        for text, cmd, color in buttons:
            tk.Button(self, text=text, bg=color, fg='white', font=('Arial', 12, 'bold'), command=cmd).pack(fill='x', pady=10, padx=20)

    def back(self):
        self.parent.show_main_frame()

    def add_user(self):
        self.parent.show_frame(AddUserFrame)

    def remove_user(self):
        self.parent.show_frame(RemoveUserFrame)

    def placeholder(self):
        messagebox.showinfo("Placeholder", "Not implemented yet.")

class AddUserFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        self.groups = []
        self.distros = []
        try:
            with parent.server.auth.sign_in(parent.auth):
                parent.server.use_highest_version()
                self.groups = sorted([g.name for g in tsc.Pager(parent.server.groups)])
            # Query Active Directory for distribution lists
            server = Server(self.parent.ldap_server, port=self.parent.ldap_port, get_info=ALL)
            conn = Connection(server, user=self.parent.ldap_user, password=self.parent.ldap_password, auto_bind=True)
            search_filter = f"(&(objectClass=group)(cn=DL-customer-service-reporting*)(managedBy=*{self.parent.ntlogin}*))"
            conn.search(self.parent.ldap_base_dn, search_filter, search_scope=SUBTREE, attributes=['cn'])
            self.distros = sorted([entry.cn.value for entry in conn.entries])
            conn.unbind()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize: {str(e)}")
            self.back()
            return

        tk.Label(self, text="User's NTLogin", bg='white').pack(pady=5)
        self.ntlogin_entry = tk.Entry(self)
        self.ntlogin_entry.pack(pady=5)

        tk.Label(self, text="Select Tableau Groups", bg='white').pack(pady=5)
        scrollbar_groups = Scrollbar(self)
        self.group_list = Listbox(self, selectmode='multiple', yscrollcommand=scrollbar_groups.set)
        scrollbar_groups.config(command=self.group_list.yview)
        for g in self.groups:
            self.group_list.insert('end', g)
        self.group_list.pack(side='left', fill='both', expand=True, pady=5, padx=10)
        scrollbar_groups.pack(side='right', fill='y')

        tk.Label(self, text="Select Distribution Lists", bg='white').pack(pady=5)
        scrollbar_distros = Scrollbar(self)
        self.distro_list = Listbox(self, selectmode='multiple', yscrollcommand=scrollbar_distros.set)
        scrollbar_distros.config(command=self.distro_list.yview)
        for d in self.distros:
            self.distro_list.insert('end', d)
        self.distro_list.pack(side='left', fill='both', expand=True, pady=5, padx=10)
        scrollbar_distros.pack(side='right', fill='y')

        tk.Button(self, text="Add", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=self.add).pack(pady=5)
        tk.Button(self, text="Back", bg='#005F9E', fg='white', font=('Arial', 12, 'bold'), command=self.back).pack(pady=5)

    def back(self):
        self.parent.show_frame(UserManagementFrame)

    def add(self):
        ntlogin = self.ntlogin_entry.get().strip()
        selected_group_indices = self.group_list.curselection()
        selected_groups = [self.groups[i] for i in selected_group_indices]
        selected_distro_indices = self.distro_list.curselection()
        selected_distros = [self.distros[i] for i in selected_distro_indices]

        if not ntlogin:
            messagebox.showerror("Error", "Enter NTLogin")
            return

        try:
            # Add user to Tableau Server
            with self.parent.server.auth.sign_in(self.parent.auth):
                new_user = tsc.UserItem(name=ntlogin, site_role='Viewer')
                new_user = self.parent.server.users.add(new_user)
                user = self.parent.server.users.get_by_id(new_user.id)
                if user.site_role == 'Unlicensed':
                    self.parent.server.users.remove(new_user.id)
                    messagebox.showerror("Error", "No licenses available to add the user.")
                    return
                for gname in selected_groups:
                    groups = [gr for gr in tsc.Pager(self.parent.server.groups) if gr.name == gname]
                    if groups:
                        self.parent.server.groups.add_user(groups[0], new_user.id)

            # Add user to Active Directory distribution lists
            if selected_distros:
                server = Server(self.parent.ldap_server, port=self.parent.ldap_port, get_info=ALL)
                conn = Connection(server, user=self.parent.ldap_user, password=self.parent.ldap_password, auto_bind=True)
                user_dn = None
                conn.search(self.parent.ldap_base_dn, f"(sAMAccountName={ntlogin})", attributes=['distinguishedName'])
                if conn.entries:
                    user_dn = conn.entries[0].distinguishedName.value
                else:
                    conn.unbind()
                    messagebox.showerror("Error", f"User {ntlogin} not found in Active Directory.")
                    return

                for distro in selected_distros:
                    conn.search(self.parent.ldap_base_dn, f"(cn={distro})", attributes=['distinguishedName'])
                    if conn.entries:
                        distro_dn = conn.entries[0].distinguishedName.value
                        conn.modify(distro_dn, {'member': [(MODIFY_ADD, [user_dn])]})
                        if not conn.result['result'] == 0:
                            conn.unbind()
                            messagebox.showerror("Error", f"Failed to add user to distribution list {distro}.")
                            return
                conn.unbind()

            messagebox.showinfo("Success", f"User {ntlogin} added to Tableau groups and {len(selected_distros)} distribution lists successfully.")
            self.back()
        except Exception as e:
            messagebox.showerror("Error", str(e))

class RemoveUserFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        self.users = []
        self.user_map = {}
        try:
            with parent.server.auth.sign_in(parent.auth):
                users = sorted([(u.name, u) for u in tsc.Pager(parent.server.users)], key=lambda x: x[0])
                for username, u in users:
                    disp = f"{u.name} - {u.fullname}"
                    self.user_map[disp] = u
                    self.users.append(disp)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.back()
            return
        tk.Label(self, text="Select User", bg='white').pack(pady=5)
        self.user_combo = ttk.Combobox(self, values=self.users)
        self.user_combo.pack(pady=5)
        tk.Button(self, text="Remove User", bg='#005F9E', fg='white', font=('Arial', 12, 'bold'), command=self.remove).pack(pady=5)
        tk.Button(self, text="Back", bg='#005F9E', fg='white', font=('Arial', 12, 'bold'), command=self.back).pack(pady=5)

    def back(self):
        self.parent.show_frame(UserManagementFrame)

    def remove(self):
        sel = self.user_combo.get()
        if not sel:
            messagebox.showerror("Error", "Select a user")
            return
        user = self.user_map[sel]
        try:
            with self.parent.server.auth.sign_in(self.parent.auth):
                self.parent.server.users.remove(user.id)
            success = tk.Toplevel(self)
            success.title("Success")
            tk.Label(success, text=f"{user.fullname} was removed from the {self.parent.site_id} site successfully.").pack(padx=20, pady=20)
            tk.Button(success, text="OK", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=lambda: [success.destroy(), self.back()]).pack(pady=10)
        except Exception as e:
            messagebox.showerror("Error", str(e))

class ReportsFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='white')
        self.parent = parent
        buttons = [
            ("User Report", self.user_report, '#005F9E'),
            ("User Group Report", self.group_report, '#005F9E'),
            ("Projects Report", self.project_report, '#005F9E'),
            ("Workbook Report", self.workbook_report, '#005F9E'),
            ("Data Source Report", self.datasource_report, '#005F9E'),
            ("Favorites Report", self.favorites_report, '#005F9E'),
            ("Subscriptions Report", self.subscriptions_report, '#005F9E'),
            ("Master Report", self.master_report, '#005F9E'),
            ("Back", self.back, '#005F9E'),
        ]
        for text, cmd, color in buttons:
            tk.Button(self, text=text, bg=color, fg='white', font=('Arial', 12, 'bold'), command=cmd).pack(fill='x', pady=10, padx=20)

    def back(self):
        self.parent.show_main_frame()

    def start_report(self, func):
        self.loading = tk.Toplevel(self)
        self.loading.title("Loading")
        tk.Label(self.loading, text="Generating report...").pack(padx=20, pady=20)
        thread = threading.Thread(target=self.run_report, args=(func,))
        thread.start()

    def run_report(self, func):
        try:
            file_path = func()
            self.parent.after(0, lambda: self.show_success(file_path))
        except Exception as exc:
            error_message = str(exc)
            self.parent.after(0, lambda: self.show_error(error_message))

    def show_success(self, path):
        self.loading.destroy()
        success = tk.Toplevel(self)
        success.title("Success")
        tk.Label(success, text=f"Report generated successfully and saved to {path}.").pack(padx=20, pady=20)
        tk.Button(success, text="OK", bg='#00BFFF', fg='white', font=('Arial', 12, 'bold'), command=success.destroy).pack(pady=10)

    def show_error(self, err):
        self.loading.destroy()
        messagebox.showerror("Error", err)

    def user_report(self):
        self.start_report(self.generate_user_report)

    def generate_user_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            users = [user for user in tsc.Pager(self.parent.server.users)]
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Name': user.name,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Domain': self.parent.server.users.get_by_id(user.id).domain_name,
                        'User Site Role': user.site_role
                    }
                    for user in users
                ]
            )
            data_source_owners = [data_source.owner_id for data_source in tsc.Pager(self.parent.server.datasources)]
            workbook_owners = [workbook.owner_id for workbook in tsc.Pager(self.parent.server.workbooks)]
            owners = set(data_source_owners)
            owners.update(workbook_owners)
            users_report = (
                user_info
                .assign(
                    is_owner=user_info['User ID'].isin(owners)
                )
                .rename(columns={'is_owner': 'Content Owner'})
                .sort_values(by='Content Owner', ascending=False)
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - User Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            users_report.to_excel(writer, sheet_name='Users', index=False)
        return path

    def group_report(self):
        self.start_report(self.generate_group_report)

    def generate_group_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            groups = [group for group in tsc.Pager(self.parent.server.groups)]
            for group in groups:
                self.parent.server.groups.populate_users(group)
            group_info = pd.DataFrame(
                [
                    {
                        'Group ID': group.id,
                        'Group Name': group.name,
                        'Group Domain': group.domain_name,
                        'Users': [user.id for user in group.users]
                    }
                    for group in groups
                ]
            )
            group_info = group_info.explode(column='Users')
            users = [user for user in tsc.Pager(self.parent.server.users)]
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Name': user.name,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Domain': self.parent.server.users.get_by_id(user.id).domain_name,
                        'User Site Role': user.site_role
                    }
                    for user in users
                ]
            )
            groups_report = (
                group_info
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Users',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - User Group Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            groups_report.to_excel(writer, sheet_name='Groups', index=False)
        return path

    def project_report(self):
        self.start_report(self.generate_project_report)

    def generate_project_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            projects = [project for project in tsc.Pager(self.parent.server.projects)]
            users = [user for user in tsc.Pager(self.parent.server.users)]
            project_info = pd.DataFrame(
                [
                    {
                        'Project ID': project.id,
                        'Project Name': project.name,
                        'Project Description': project.description,
                        'Project Owner ID': project.owner_id,
                        'Parent Project ID': project.parent_id
                    }
                    for project in projects
                ]
            )
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Site Role': user.site_role
                    }
                    for user in users
                ]
            )
            projects_report = (
                project_info
                .merge(
                    right=project_info,
                    how='left',
                    left_on='Parent Project ID',
                    right_on='Project ID',
                    suffixes=('', ' Parent'),
                )
                .drop(
                    columns=[
                        'Parent Project ID Parent',
                        'Project ID Parent'
                    ]
                )
                .rename(
                    columns={
                        'Project Name Parent': 'Parent Project Name',
                        'Project Description Parent': 'Parent Project Description',
                        'Project Owner ID Parent': 'Parent Project Owner ID'
                    }
                )
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Project Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Project Owner Name',
                        'User Email Address': 'Project Owner Email Address',
                        'User Site Role': 'Project Owner Site Role'
                    }
                )
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Parent Project Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Parent Project Owner Name',
                        'User Email Address': 'Parent Project Owner Email Address',
                        'User Site Role': 'Parent Project Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Project ID', 'Project Name', 'Project Description',
                        'Project Owner ID', 'Project Owner Name', 'Project Owner Email Address',
                        'Project Owner Site Role', 'Parent Project ID', 'Parent Project Name',
                        'Parent Project Description', 'Parent Project Owner ID', 'Parent Project Owner Name',
                        'Parent Project Owner Email Address', 'Parent Project Owner Site Role'
                    ]
                )
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Projects Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            projects_report.to_excel(writer, sheet_name='Projects', index=False)
        return path

    def workbook_report(self):
        self.start_report(self.generate_workbook_report)

    def generate_workbook_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            workbooks = [workbook for workbook in tsc.Pager(self.parent.server.workbooks)]
            views = [view for view in tsc.Pager(self.parent.server.views, usage=True)]
            users = [user for user in tsc.Pager(self.parent.server.users)]
            tasks = [task for task in tsc.Pager(self.parent.server.tasks)]
            refresh_tasks = [
                {
                    'Workbook ID': task.target.id,
                    'Last Refresh Duration (Seconds)': (task.completed_at - task.started_at).total_seconds()
                    if task.completed_at and task.started_at else None
                }
                for task in tasks
                if task.task_type == 'refresh_extract' and task.target.type == 'workbook' and task.completed_at
            ]
            refresh_df = pd.DataFrame(refresh_tasks)
            latest_refresh = (
                refresh_df
                .groupby('Workbook ID')
                ['Last Refresh Duration (Seconds)']
                .last()
                .reset_index()
            )
            workbook_info = pd.DataFrame(
                [
                    {
                        'Workbook ID': workbook.id,
                        'Workbook Owner ID': workbook.owner_id,
                        'Workbook Name': workbook.name,
                        'Workbook Created At': workbook.created_at,
                        'Workbook Updated At': workbook.updated_at,
                        'Workbook Content URL': workbook.webpage_url,
                        'Workbook Project ID': workbook.project_id,
                        'Workbook Project Name': workbook.project_name,
                        'Workbook Size (Bytes)': workbook.size
                    }
                    for workbook in workbooks
                ]
            )
            total_views_per_workbook = (
                pd.DataFrame([view.__dict__ for view in views])
                .groupby('_workbook_id')['_total_views']
                .sum()
                .reset_index()
                .rename(
                    columns={
                        '_workbook_id': 'Workbook ID',
                        '_total_views': 'Workbook Total Views'
                    }
                )
            )
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Site Role': user.site_role,
                    }
                    for user in users
                ]
            )
            workbooks_report = (
                workbook_info
                .merge(
                    right=total_views_per_workbook,
                    how='inner',
                    on='Workbook ID'
                )
                .merge(
                    right=user_info,
                    how='inner',
                    left_on='Workbook Owner ID',
                    right_on='User ID'
                )
                .merge(
                    right=latest_refresh,
                    how='left',
                    on='Workbook ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Workbook Owner Name',
                        'User Email Address': 'Workbook Owner Email Address',
                        'User Site Role': 'Workbook Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Workbook ID', 'Workbook Name',
                        'Workbook Total Views', 'Workbook Created At',
                        'Workbook Updated At', 'Workbook Content URL',
                        'Workbook Project ID', 'Workbook Project Name',
                        'Workbook Owner ID', 'Workbook Owner Name',
                        'Workbook Owner Email Address', 'Workbook Owner Site Role',
                        'Workbook Size (Bytes)', 'Last Refresh Duration (Seconds)'
                    ]
                )
            )
            workbooks_report['Workbook Created At'] = workbooks_report['Workbook Created At'].dt.tz_localize(None)
            workbooks_report['Workbook Updated At'] = workbooks_report['Workbook Updated At'].dt.tz_localize(None)
            workbooks_report['Last Refresh Duration (Seconds)'] = workbooks_report['Last Refresh Duration (Seconds)'].fillna('N/A')
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Workbook Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            workbooks_report.to_excel(writer, sheet_name='Workbooks', index=False)
        return path

    def datasource_report(self):
        self.start_report(self.generate_datasource_report)

    def generate_datasource_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            data_sources = [data_source for data_source in tsc.Pager(self.parent.server.datasources)]
            users = [user for user in tsc.Pager(self.parent.server.users)]
            data_source_info = pd.DataFrame(
                [
                    {
                        'Data Source ID': data_source.id,
                        'Data Source Owner ID': data_source.owner_id,
                        'Data Source Name': data_source.name,
                        'Data Source Type': data_source.datasource_type,
                        'Data Source Created At': data_source.created_at,
                        'Data Source Updated At': data_source.updated_at,
                        'Data Source Project ID': data_source.project_id,
                        'Data Source Project Name': data_source.project_name,
                    }
                    for data_source in data_sources
                ]
            )
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Site Role': user.site_role,
                    }
                    for user in users
                ]
            )
            data_sources_report = (
                data_source_info
                .merge(
                    right=user_info,
                    how='inner',
                    left_on='Data Source Owner ID',
                    right_on='User ID',
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Data Source Owner Name',
                        'User Email Address': 'Data Source Owner Email Address',
                        'User Site Role': 'Data Source Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Data Source ID', 'Data Source Name',
                        'Data Source Type', 'Data Source Created At',
                        'Data Source Updated At', 'Data Source Project ID',
                        'Data Source Project Name', 'Data Source Owner ID',
                        'Data Source Owner Name', 'Data Source Owner Email Address',
                        'Data Source Owner Site Role'
                    ]
                )
            )
            data_sources_report['Data Source Created At'] = data_sources_report['Data Source Created At'].dt.tz_localize(None)
            data_sources_report['Data Source Updated At'] = data_sources_report['Data Source Updated At'].dt.tz_localize(None)
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Data Source Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            data_sources_report.to_excel(writer, sheet_name='Data Sources', index=False)
        return path

    def favorites_report(self):
        self.start_report(self.generate_favorites_report)

    def generate_favorites_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            users = [user for user in tsc.Pager(self.parent.server.users)]
            all_favorites = []
            for user in users:
                self.parent.server.users.populate_favorites(user)
                favorite_categories = [category for category in user.favorites.keys() if user.favorites[category]]
                for category in favorite_categories:
                    category_favorites = pd.DataFrame(data=user.favorites[category], columns=['Favorite'])
                    category_favorites = (
                        category_favorites
                        .assign(
                            favorite_id=category_favorites['Favorite'].apply(lambda favorite: favorite.id),
                            favorite_name=category_favorites['Favorite'].apply(lambda favorite: favorite.name),
                            favorite_category=category.title(),
                            favorite_project_id=category_favorites['Favorite'].apply(
                                lambda favorite: getattr(favorite, 'project_id', 'Not applicable')
                            ),
                            favorite_project_name=category_favorites['Favorite'].apply(
                                lambda favorite: getattr(favorite, 'project_name', 'Not applicable')
                            ),
                            user_id=user.id,
                            user_display_name=user.fullname,
                            user_email_address=user.email,
                            user_site_role=user.site_role
                        )
                    )
                    all_favorites.append(category_favorites)
            favorites_report = (
                pd.concat(
                    all_favorites,
                    axis=0,
                    ignore_index=True
                )
                .drop(columns=['Favorite'])
                .rename(
                    lambda column: column.replace('_', ' ').title().replace('Id', 'ID'),
                    axis=1
                )
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Favorites Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            favorites_report.to_excel(writer, sheet_name='Favorites', index=False)
        return path

    def subscriptions_report(self):
        self.start_report(self.generate_subscriptions_report)

    def generate_subscriptions_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            subscriptions = [subscription for subscription in tsc.Pager(self.parent.server.subscriptions)]
            workbooks = [workbook for workbook in tsc.Pager(self.parent.server.workbooks)]
            views = [view for view in tsc.Pager(self.parent.server.views)]
            users = [user for user in tsc.Pager(self.parent.server.users)]
            subscription_info = pd.DataFrame(
                [
                    {
                        'Subscription ID': subscription.id,
                        'Subscription Owner ID': subscription.user_id,
                        'Subscription Subject': subscription.subject,
                        'Subscription Content ID': subscription.target.id,
                        'Subscription Content Type': subscription.target.type,
                        'Subscription Schedule': subscription.schedule[0].interval_item if subscription.schedule and len(subscription.schedule) > 0 else 'N/A'
                    }
                    for subscription in subscriptions
                ]
            )
            workbook_info = pd.DataFrame(
                [
                    {
                        'Content ID': workbook.id,
                        'Content Owner ID': workbook.owner_id,
                        'Content Name': workbook.name,
                        'Content URL': workbook.webpage_url
                    }
                    for workbook in workbooks
                ]
            )
            view_info = pd.DataFrame(
                [
                    {
                        'Content ID': view.id,
                        'Content Owner ID': view.owner_id,
                        'Content Name': view.name,
                        'Content URL': 
                            self.parent.server.server_address
                            + '/#/site/' + self.parent.site_id + '/views/'
                            + view.content_url.replace('/sheets/', '/') 
                    }
                    for view in views
                ]
            )
            content_info = pd.concat(
                [workbook_info, view_info],
                axis=0,
                ignore_index=True
            )
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Site Role': user.site_role
                    }
                    for user in users
                ]
            )
            subscriptions_report = (
                subscription_info
                .merge(
                    right=content_info,
                    left_on='Subscription Content ID',
                    right_on='Content ID'
                )
                .drop(columns=['Subscription Content ID'])
                .merge(
                    right=user_info,
                    left_on='Subscription Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Subscription Owner Display Name',
                        'User Email Address': 'Subscription Owner Email Address',
                        'User Site Role': 'Subscription Owner Site Role'
                    }
                )
                .merge(
                    right=user_info,
                    left_on='Content Owner ID',
                    right_on='User ID'       
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Content Owner Display Name',
                        'User Email Address': 'Content Owner Email Address',
                        'User Site Role': 'Content Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Subscription ID', 'Subscription Owner ID', 'Subscription Subject',
                        'Subscription Schedule', 'Subscription Owner Display Name', 'Subscription Owner Email Address',
                        'Subscription Owner Site Role', 'Subscription Content Type', 'Content ID',
                        'Content Name', 'Content URL', 'Content Owner ID',
                        'Content Owner Display Name', 'Content Owner Email Address', 'Content Owner Site Role'
                    ]
                )
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Subscription Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            subscriptions_report.to_excel(writer, sheet_name='Subscriptions', index=False)
        return path

    def master_report(self):
        self.start_report(self.generate_master_report)

    def generate_master_report(self):
        with self.parent.server.auth.sign_in(self.parent.auth):
            self.parent.server.use_highest_version()
            users = [user for user in tsc.Pager(self.parent.server.users)]
            for user in users:
                self.parent.server.users.populate_favorites(user)
            groups = [group for group in tsc.Pager(self.parent.server.groups)]
            for group in groups:
                self.parent.server.groups.populate_users(group)
            group_info = pd.DataFrame(
                [
                    {
                        'Group ID': group.id,
                        'Group Name': group.name,
                        'Group Domain': group.domain_name,
                        'Users': [user.id for user in group.users]
                    }
                    for group in groups
                ]
            )
            projects = [project for project in tsc.Pager(self.parent.server.projects)]
            data_sources = [data_source for data_source in tsc.Pager(self.parent.server.datasources)]
            workbooks = [workbook for workbook in tsc.Pager(self.parent.server.workbooks)]
            views = [view for view in tsc.Pager(self.parent.server.views, usage=True)]
            flows = [flow for flow in tsc.Pager(self.parent.server.flows)]
            runs = [run for run in tsc.Pager(self.parent.server.flow_runs)]
            subscriptions = [subscription for subscription in tsc.Pager(self.parent.server.subscriptions)]
            tasks = [task for task in tsc.Pager(self.parent.server.tasks)]
            user_info = pd.DataFrame(
                [
                    {
                        'User ID': user.id,
                        'User Display Name': user.fullname,
                        'User Email Address': user.email,
                        'User Site Role': user.site_role
                    }
                    for user in users
                ]
            )
            flow_owners = [flow.owner_id for flow in flows]
            data_source_owners = [data_source.owner_id for data_source in tsc.Pager(self.parent.server.datasources)]
            workbook_owners = [workbook.owner_id for workbook in tsc.Pager(self.parent.server.workbooks)]
            flow_owners = set(flow_owners)
            data_source_owners = set(data_source_owners)
            workbook_owners = set(workbook_owners)
            flow_owners.update(data_source_owners)
            flow_owners.update(workbook_owners)
            owners = set(flow_owners)
            users_report = (
                user_info
                .assign(
                    is_owner=user_info['User ID'].isin(owners)
                )
                .rename(columns={'is_owner': 'Content Owner'})
                .sort_values(by='Content Owner', ascending=False)
            )
            all_favorites = []
            for user in users:
                favorite_categories = [category for category in user.favorites.keys() if user.favorites[category]]
                for category in favorite_categories:
                    category_favorites = pd.DataFrame(data=user.favorites[category], columns=['Favorite'])
                    category_favorites = (
                        category_favorites
                        .assign(
                            favorite_id=category_favorites['Favorite'].apply(lambda favorite: favorite.id),
                            favorite_name=category_favorites['Favorite'].apply(lambda favorite: favorite.name),
                            favorite_category=category.title(),
                            favorite_project_id=category_favorites['Favorite'].apply(
                                lambda favorite: getattr(favorite, 'project_id', 'Not applicable')
                            ),
                            favorite_project_name=category_favorites['Favorite'].apply(
                                lambda favorite: getattr(favorite, 'project_name', 'Not applicable')
                            ),
                            user_id=user.id,
                            user_display_name=user.fullname,
                            user_email_address=user.email,
                            user_site_role=user.site_role
                        )
                    )
                    all_favorites.append(category_favorites)
            favorites_report = (
                pd.concat(
                    all_favorites,
                    axis=0,
                    ignore_index=True
                )
                .drop(columns=['Favorite'])
                .rename(
                    lambda column: column.replace('_', ' ').title().replace('Id', 'ID'),
                    axis=1
                )
            )
            group_info = group_info.explode(column='Users')
            groups_report = (
                group_info
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Users',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
            )
            project_info = pd.DataFrame(
                [
                    {
                        'Project ID': project.id,
                        'Project Name': project.name,
                        'Project Description': project.description,
                        'Project Owner ID': project.owner_id,
                        'Parent Project ID': project.parent_id
                    }
                    for project in projects
                ]
            )
            projects_report = (
                project_info
                .merge(
                    right=project_info,
                    how='left',
                    left_on='Parent Project ID',
                    right_on='Project ID',
                    suffixes=('', ' Parent'),
                )
                .drop(
                    columns=[
                        'Parent Project ID Parent',
                        'Project ID Parent'
                    ]
                )
                .rename(
                    columns={
                        'Project Name Parent': 'Parent Project Name',
                        'Project Description Parent': 'Parent Project Description',
                        'Project Owner ID Parent': 'Parent Project Owner ID'
                    }
                )
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Project Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Project Owner Name',
                        'User Email Address': 'Project Owner Email Address',
                        'User Site Role': 'Project Owner Site Role'
                    }
                )
                .merge(
                    right=user_info,
                    how='left',
                    left_on='Parent Project Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Parent Project Owner Name',
                        'User Email Address': 'Parent Project Owner Email Address',
                        'User Site Role': 'Parent Project Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Project ID', 'Project Name', 'Project Description',
                        'Project Owner ID', 'Project Owner Name', 'Project Owner Email Address',
                        'Project Owner Site Role', 'Parent Project ID', 'Parent Project Name',
                        'Parent Project Description', 'Parent Project Owner ID', 'Parent Project Owner Name',
                        'Parent Project Owner Email Address', 'Parent Project Owner Site Role'
                    ]
                )
            )
            data_source_info = pd.DataFrame(
                [
                    {
                        'Data Source ID': data_source.id,
                        'Data Source Owner ID': data_source.owner_id,
                        'Data Source Name': data_source.name,
                        'Data Source Type': data_source.datasource_type,
                        'Data Source Created At': data_source.created_at,
                        'Data Source Updated At': data_source.updated_at,
                        'Data Source Project ID': data_source.project_id,
                        'Data Source Project Name': data_source.project_name,
                    }
                    for data_source in data_sources
                ]
            )
            data_sources_report = (
                data_source_info
                .merge(
                    right=user_info,
                    how='inner',
                    left_on='Data Source Owner ID',
                    right_on='User ID',
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Data Source Owner Name',
                        'User Email Address': 'Data Source Owner Email Address',
                        'User Site Role': 'Data Source Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Data Source ID', 'Data Source Name',
                        'Data Source Type', 'Data Source Created At',
                        'Data Source Updated At', 'Data Source Project ID',
                        'Data Source Project Name', 'Data Source Owner ID',
                        'Data Source Owner Name', 'Data Source Owner Email Address',
                        'Data Source Owner Site Role'
                    ]
                )
            )
            data_sources_report['Data Source Created At'] = data_sources_report['Data Source Created At'].dt.tz_localize(None)
            data_sources_report['Data Source Updated At'] = data_sources_report['Data Source Updated At'].dt.tz_localize(None)
            refresh_tasks = [
                {
                    'Workbook ID': task.target.id,
                    'Last Refresh Duration (Seconds)': (task.completed_at - task.started_at).total_seconds()
                    if task.completed_at and task.started_at else None
                }
                for task in tasks
                if task.task_type == 'refresh_extract' and task.target.type == 'workbook' and task.completed_at
            ]
            refresh_df = pd.DataFrame(refresh_tasks)
            latest_refresh = (
                refresh_df
                .groupby('Workbook ID')
                ['Last Refresh Duration (Seconds)']
                .last()
                .reset_index()
            )
            workbook_info = pd.DataFrame(
                [
                    {
                        'Workbook ID': workbook.id,
                        'Workbook Owner ID': workbook.owner_id,
                        'Workbook Name': workbook.name,
                        'Workbook Created At': workbook.created_at,
                        'Workbook Updated At': workbook.updated_at,
                        'Workbook Content URL': workbook.webpage_url,
                        'Workbook Project ID': workbook.project_id,
                        'Workbook Project Name': workbook.project_name,
                        'Workbook Size (Bytes)': workbook.size
                    }
                    for workbook in workbooks
                ]
            )
            total_views_per_workbook = (
                pd.DataFrame([view.__dict__ for view in views])
                .groupby('_workbook_id')['_total_views']
                .sum()
                .reset_index()
                .rename(
                    columns={
                        '_workbook_id': 'Workbook ID',
                        '_total_views': 'Workbook Total Views'
                    }
                )
            )
            workbooks_report = (
                workbook_info
                .merge(
                    right=total_views_per_workbook,
                    how='inner',
                    on='Workbook ID'
                )
                .merge(
                    right=user_info,
                    how='inner',
                    left_on='Workbook Owner ID',
                    right_on='User ID'
                )
                .merge(
                    right=latest_refresh,
                    how='left',
                    on='Workbook ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Workbook Owner Name',
                        'User Email Address': 'Workbook Owner Email Address',
                        'User Site Role': 'Workbook Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Workbook ID', 'Workbook Name',
                        'Workbook Total Views', 'Workbook Created At',
                        'Workbook Updated At', 'Workbook Content URL',
                        'Workbook Project ID', 'Workbook Project Name',
                        'Workbook Owner ID', 'Workbook Owner Name',
                        'Workbook Owner Email Address', 'Workbook Owner Site Role',
                        'Workbook Size (Bytes)', 'Last Refresh Duration (Seconds)'
                    ]
                )
            )
            workbooks_report['Workbook Created At'] = workbooks_report['Workbook Created At'].dt.tz_localize(None)
            workbooks_report['Workbook Updated At'] = workbooks_report['Workbook Updated At'].dt.tz_localize(None)
            workbooks_report['Last Refresh Duration (Seconds)'] = workbooks_report['Last Refresh Duration (Seconds)'].fillna('N/A')
            flow_info = pd.DataFrame(
                [
                    {
                        'Flow ID': flow.id,
                        'Flow Owner ID': flow.owner_id,
                        'Flow Name': flow.name,
                        'Flow Project ID': flow.project_id,
                        'Flow Project Name': flow.project_name,
                        'Flow Content URL': flow.webpage_url
                    }
                    for flow in flows
                ]
            )
            flow_run_history = pd.DataFrame(
                [
                    {
                        'Flow ID': run.flow_id,
                        'Run Duration': run.completed_at - run.started_at
                    }
                    for run in runs
                ]
            )
            flow_run_summary = (
                flow_run_history
                .groupby('Flow ID')['Run Duration']
                .agg(['count', 'sum', 'mean', 'max', 'min'])
                .assign(duration_range=lambda flow: flow['max'] - flow['min'])
                .reset_index()
                .rename(
                    columns={
                        'count': 'Run Count',
                        'sum': 'Total Duration',
                        'mean': 'Average Duration',
                        'max': 'Maximum Duration',
                        'min': 'Minimum Duration',
                        'duration_range': 'Duration Range'
                    }
                )
            )
            flow_run_summary['Total Duration'] = flow_run_summary['Total Duration'].dt.total_seconds()
            flow_run_summary['Average Duration'] = flow_run_summary['Average Duration'].dt.total_seconds()
            flow_run_summary['Maximum Duration'] = flow_run_summary['Maximum Duration'].dt.total_seconds()
            flow_run_summary['Minimum Duration'] = flow_run_summary['Minimum Duration'].dt.total_seconds()
            flow_run_summary['Duration Range'] = flow_run_summary['Duration Range'].dt.total_seconds()
            flows_report = (
                flow_info
                .merge(
                    right=user_info,
                    how='inner',
                    left_on='Flow Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .merge(
                    right=flow_run_summary,
                    how='inner',
                    on='Flow ID'
                )
            )
            subscription_info = pd.DataFrame(
                [
                    {
                        'Subscription ID': subscription.id,
                        'Subscription Owner ID': subscription.user_id,
                        'Subscription Subject': subscription.subject,
                        'Subscription Content ID': subscription.target.id,
                        'Subscription Content Type': subscription.target.type,
                        'Subscription Schedule': subscription.schedule[0].interval_item if subscription.schedule and len(subscription.schedule) > 0 else 'N/A'
                    }
                    for subscription in subscriptions
                ]
            )
            view_info = pd.DataFrame(
                [
                    {
                        'Content ID': view.id,
                        'Content Owner ID': view.owner_id,
                        'Content Name': view.name,
                        'Content URL': 
                            self.parent.server.server_address
                            + '/#/site/' + self.parent.site_id + '/views/'
                            + view.content_url.replace('/sheets/', '/') 
                    }
                    for view in views
                ]
            )
            content_info = pd.concat(
                [
                    workbook_info
                    .loc[:, ['Workbook ID', 'Workbook Owner ID', 'Workbook Name', 'Workbook Content URL']]
                    .rename(
                        columns={
                            'Workbook ID': 'Content ID',
                            'Workbook Owner ID': 'Content Owner ID',
                            'Workbook Name': 'Content Name',
                            'Workbook Content URL': 'Content URL'
                        }
                    ),
                    view_info
                ],
                axis=0,
                ignore_index=True
            )
            subscriptions_report = (
                subscription_info
                .merge(
                    right=content_info,
                    left_on='Subscription Content ID',
                    right_on='Content ID'
                )
                .drop(columns=['Subscription Content ID'])
                .merge(
                    right=user_info,
                    left_on='Subscription Owner ID',
                    right_on='User ID'
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Subscription Owner Display Name',
                        'User Email Address': 'Subscription Owner Email Address',
                        'User Site Role': 'Subscription Owner Site Role'
                    }
                )
                .merge(
                    right=user_info,
                    left_on='Content Owner ID',
                    right_on='User ID'       
                )
                .drop(columns=['User ID'])
                .rename(
                    columns={
                        'User Display Name': 'Content Owner Display Name',
                        'User Email Address': 'Content Owner Email Address',
                        'User Site Role': 'Content Owner Site Role'
                    }
                )
                .reindex(
                    columns=[
                        'Subscription ID', 'Subscription Owner ID', 'Subscription Subject',
                        'Subscription Schedule', 'Subscription Owner Display Name', 'Subscription Owner Email Address',
                        'Subscription Owner Site Role', 'Subscription Content Type', 'Content ID',
                        'Content Name', 'Content URL', 'Content Owner ID',
                        'Content Owner Display Name', 'Content Owner Email Address', 'Content Owner Site Role'
                    ]
                )
            )
        documents = os.path.join(os.environ['USERPROFILE'], 'Documents')
        path = os.path.join(documents, 'Tableau Server - Master Report.xlsx')
        with pd.ExcelWriter(path) as writer:
            users_report.to_excel(writer, sheet_name='Users', index=False)
            favorites_report.to_excel(writer, sheet_name='Favorites', index=False)
            groups_report.to_excel(writer, sheet_name='Groups', index=False)
            projects_report.to_excel(writer, sheet_name='Projects', index=False)
            data_sources_report.to_excel(writer, sheet_name='Data Sources', index=False)
            workbooks_report.to_excel(writer, sheet_name='Workbooks', index=False)
            flows_report.to_excel(writer, sheet_name='Flows', index=False)
            subscriptions_report.to_excel(writer, sheet_name='Subscriptions', index=False)
        return path

if __name__ == '__main__':
    app = App()
    app.mainloop()