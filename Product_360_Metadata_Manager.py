import os
import shutil
from datetime import date
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import logging
import time

# Set up logging to a file on the desktop
log_file = os.path.join(os.path.expandvars('%userprofile%'), 'Desktop', 'p360_layout_manager.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Define archive folder base path
archive_folder = os.path.join(os.path.expandvars('%userprofile%'), 'Desktop', 'Product 360 layout backup')

# Define source files relative to source_base
source_files_relative = {
    'org.eclipse.ui.workbench.prefs': os.path.join('.metadata', '.plugins', 'org.eclipse.core.runtime', '.settings', 'org.eclipse.ui.workbench.prefs'),
    'workbench.xmi': os.path.join('.metadata', '.plugins', 'org.eclipse.e4.workbench', 'workbench.xmi'),
    'savedTableConfigs.xml': os.path.join('.metadata', '.plugins', 'com.heiler.ppm.std.ui', 'savedTableConfigs.xml')
}

# Folders to exclude (case-insensitive)
EXCLUDED_FOLDERS = {
    '$recycle.bin', 'system volume information', 'windows', 'program files', 
    'program files (x86)', 'onedrive', 'temp', '$windows.~bt'
}

# Likely folder keywords (case-insensitive)
LIKELY_KEYWORDS = ['informatica', 'pim', 'product 360']

def is_excluded_folder(path):
    """Check if a folder should be skipped based on its name or path."""
    path_lower = path.lower()
    folder_name = os.path.basename(path_lower)
    return (folder_name in EXCLUDED_FOLDERS or 
            any(excluded in path_lower for excluded in EXCLUDED_FOLDERS))

def is_likely_folder(path):
    """Check if a folder is a candidate for Product 360 based on its name."""
    path_lower = path.lower()
    return any(keyword in path_lower for keyword in LIKELY_KEYWORDS)

def find_install_folders(root_dir, depth=3, prioritize_likely=True):
    """Recursively search for folders containing pim-desktop.cmd up to a specified depth."""
    install_folders = []
    try:
        for item in os.listdir(root_dir):
            path = os.path.join(root_dir, item)
            if os.path.isdir(path) and not is_excluded_folder(path):
                logging.debug(f"Checking: {path}")
                cmd_path = os.path.join(path, 'pim-desktop.cmd')
                if os.path.exists(cmd_path):
                    install_folders.append(path)
                    logging.info(f"Found pim-desktop.cmd in {path}")
                elif depth > 0:
                    if prioritize_likely and is_likely_folder(path):
                        sub_folders = find_install_folders(path, depth - 1, prioritize_likely)
                        install_folders.extend(sub_folders)
                    elif not prioritize_likely:
                        sub_folders = find_install_folders(path, depth - 1, prioritize_likely)
                        install_folders.extend(sub_folders)
    except Exception as e:
        logging.error(f"Error in {root_dir}: {str(e)}")
    return install_folders

def autodetect_environments():
    """Detect Product 360 installations by searching likely folders and user directories."""
    root_dirs = ['C:\\', os.path.expandvars('%userprofile%')]  # Search C:\ and user profile
    logging.info(f"Starting search in {root_dirs}")
    install_folders = []
    for root_dir in root_dirs:
        logging.info(f"Searching likely folders in {root_dir}")
        likely_folders = find_install_folders(root_dir, depth=3, prioritize_likely=True)
        install_folders.extend(likely_folders)
        logging.info(f"Searching other folders in {root_dir} with depth=1")
        other_folders = find_install_folders(root_dir, depth=1, prioritize_likely=False)
        install_folders.extend(other_folders)
    if not install_folders:
        logging.info("No installations found")
        messagebox.showinfo("Autodetect", "No environments found.")
    return install_folders

def get_workspace_dir(install_folder):
    """Extract WORKSPACE_DIR from pim-desktop.cmd in the install folder."""
    cmd_path = os.path.join(install_folder, 'pim-desktop.cmd')
    if not os.path.exists(cmd_path):
        logging.warning(f"pim-desktop.cmd not found in {install_folder}")
        return None
    try:
        with open(cmd_path, 'r') as f:
            for line in f:
                if line.strip().startswith('SET WORKSPACE_DIR='):
                    workspace_dir = line.strip().split('=', 1)[1]
                    if os.path.exists(workspace_dir):
                        logging.info(f"Found WORKSPACE_DIR: {workspace_dir}")
                        return workspace_dir
                    else:
                        logging.warning(f"WORKSPACE_DIR {workspace_dir} does not exist")
                        return None
        logging.warning(f"SET WORKSPACE_DIR not found in {cmd_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading {cmd_path}: {str(e)}")
        return None

def infer_environment(workspace_dir):
    """Infer environment name from WORKSPACE_DIR path."""
    path_lower = workspace_dir.lower()
    if 'prod' in path_lower:
        return 'Prod'
    elif 'qa' in path_lower or 'pimqa' in path_lower:
        return 'QA'
    elif 'dev' in path_lower or 'pimdev' in path_lower:
        return 'Dev'
    else:
        return 'Custom'

def add_environment_manually():
    """Manually add an environment by selecting an install folder."""
    folder = filedialog.askdirectory(title="Select Product 360 Install Folder")
    if folder:
        workspace_dir = get_workspace_dir(folder)
        if workspace_dir:
            env = infer_environment(workspace_dir)
            logging.info(f"Manually added {env}: {workspace_dir}")
            return env, workspace_dir
        else:
            messagebox.showerror("Error", "Could not find or read WORKSPACE_DIR in pim-desktop.cmd")
    return None, None

def backup_files(env, environments, date_dropdown, selected_date):
    """Backup files for the selected environment."""
    try:
        if not env:
            messagebox.showerror("Error", "Please select an environment.")
            return
        source_base = environments[env]
        today = date.today().isoformat()
        dated_folder = os.path.join(archive_folder, env, today)
        os.makedirs(dated_folder, exist_ok=True)
        logging.info(f"Backing up to {dated_folder}")

        for filename, rel_path in source_files_relative.items():
            source_path = os.path.join(source_base, rel_path)
            if os.path.exists(source_path):
                dest_file = os.path.join(dated_folder, filename)
                shutil.copy2(source_path, dest_file)
                logging.info(f"Copied {filename} to {dest_file}")
            else:
                logging.warning(f"Source file not found: {source_path}")

        update_dates(env, environments, date_dropdown, selected_date)
        messagebox.showinfo("Backup Complete", f"Files backed up to:\n{dated_folder}")
    except Exception as e:
        logging.error(f"Backup error: {str(e)}")
        messagebox.showerror("Backup Error", f"An error occurred:\n{str(e)}")

def restore_files(env, selected_date, environments):
    """Restore files for the selected environment from the chosen date."""
    try:
        if not env or not selected_date:
            messagebox.showerror("Error", "Please select an environment and a date.")
            return
        source_base = environments[env]
        restore_folder = os.path.join(archive_folder, env, selected_date)
        if not os.path.exists(restore_folder):
            messagebox.showerror("Restore Error", f"No backup found for {selected_date} in {env}")
            return
        logging.info(f"Restoring from {restore_folder}")

        for filename, rel_path in source_files_relative.items():
            dest_path = os.path.join(source_base, rel_path)
            source_file = os.path.join(restore_folder, filename)
            if os.path.exists(source_file):
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                shutil.copy2(source_file, dest_path)
                logging.info(f"Restored {filename} to {dest_path}")
            else:
                logging.warning(f"File not found in backup: {source_file}")

        messagebox.showinfo("Restore Complete", f"Files restored from:\n{restore_folder}")
    except Exception as e:
        logging.error(f"Restore error: {str(e)}")
        messagebox.showerror("Restore Error", f"An error occurred:\n{str(e)}")

def clear_metadata(env, environments):
    """Clear the .metadata folder for the selected environment."""
    try:
        if not env:
            messagebox.showerror("Error", "Please select an environment.")
            return
        source_base = environments[env]
        metadata_folder = os.path.join(source_base, '.metadata')
        if not os.path.exists(metadata_folder):
            messagebox.showinfo("Clear Metadata", "No .metadata folder found.")
            return
        response = messagebox.askyesno(
            "Confirm Clear Metadata",
            f"Are you sure you want to delete the .metadata folder for {env}?\nThis action cannot be undone.",
            icon='warning'
        )
        if response:
            shutil.rmtree(metadata_folder)
            logging.info(f"Cleared metadata folder: {metadata_folder}")
            messagebox.showinfo("Clear Complete", f"The .metadata folder for {env} has been deleted.")
    except Exception as e:
        logging.error(f"Clear metadata error: {str(e)}")
        messagebox.showerror("Clear Error", f"An error occurred:\n{str(e)}")

def update_dates(env, environments, date_dropdown, selected_date):
    """Update the date dropdown based on the selected environment."""
    if env:
        env_archive = os.path.join(archive_folder, env)
        if os.path.exists(env_archive):
            dates = [f for f in os.listdir(env_archive) if os.path.isdir(os.path.join(env_archive, f))]
            dates.sort(reverse=True)
            date_dropdown['values'] = dates
            if dates:
                selected_date.set(dates[0])
            else:
                selected_date.set("No backups found")
        else:
            date_dropdown['values'] = []
            selected_date.set("No backups found")
    else:
        date_dropdown['values'] = []
        selected_date.set("Select an environment")

def create_gui(environments):
    """Create the main GUI."""
    root = tk.Tk()
    root.title("Product 360 Layout Manager")
    root.geometry("400x300")

    # Environment dropdown
    tk.Label(root, text="Select Environment:").pack(pady=5)
    selected_env = tk.StringVar(root)
    env_dropdown = ttk.Combobox(root, textvariable=selected_env, values=list(environments.keys()), state="readonly")
    env_dropdown.pack(pady=5)

    # Date dropdown
    tk.Label(root, text="Restore from Backup:").pack(pady=5)
    selected_date = tk.StringVar(root)
    date_dropdown = ttk.Combobox(root, textvariable=selected_date, state="readonly")
    date_dropdown.pack(pady=5)

    # Bind environment selection to update dates
    def on_env_select(*args):
        update_dates(selected_env.get(), environments, date_dropdown, selected_date)
    selected_env.trace('w', on_env_select)

    # Buttons
    tk.Button(root, text="Backup Now",
              command=lambda: backup_files(selected_env.get(), environments, date_dropdown, selected_date)).pack(pady=10)
    tk.Button(root, text="Restore Selected",
              command=lambda: restore_files(selected_env.get(), selected_date.get(), environments)).pack(pady=10)
    tk.Button(root, text="Clear Metadata",
              command=lambda: clear_metadata(selected_env.get(), environments)).pack(pady=10)

    # Initial setup
    if environments:
        selected_env.set(list(environments.keys())[0])
        update_dates(selected_env.get(), environments, date_dropdown, selected_date)

    root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide main window initially
    logging.info("Application started")

    # Ask user how to define environments
    choice = messagebox.askyesno("Define Environments", "Would you like to manually define the install folder? (No = Autodetect)")
    environments = {}

    if choice:  # Manual definition
        while True:
            env, workspace_dir = add_environment_manually()
            if env and workspace_dir:
                environments[env] = workspace_dir
                more = messagebox.askyesno("More Environments", "Add another environment?")
                if not more:
                    break
            else:
                break
    else:  # Autodetect
        install_folders = autodetect_environments()
        for folder in install_folders:
            workspace_dir = get_workspace_dir(folder)
            if workspace_dir:
                env = infer_environment(workspace_dir)
                environments[env] = workspace_dir

    if not environments:
        messagebox.showerror("Error", "No environments defined. Exiting.")
        logging.error("No environments defined")
        root.destroy()
    else:
        root.deiconify()
        create_gui(environments)