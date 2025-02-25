import os
import shutil
import json
import logging
from datetime import date
from tkinter import ttk, messagebox, Tk, StringVar
from tkinter.filedialog import askdirectory
import traceback

# Constants
ARCHIVE_BASE = os.path.join(os.path.expanduser("~/Desktop"), "Product 360 layout backup")
CONFIG_FILE = os.path.join(ARCHIVE_BASE, "environments.json")
LOG_FILE = os.path.join(ARCHIVE_BASE, "layout_manager.log")

SOURCE_FILES = {
    "workbench.xmi": os.path.join(".metadata", ".plugins", "org.eclipse.e4.workbench", "workbench.xmi"),
    "org.eclipse.ui.workbench.prefs": os.path.join(".metadata", ".plugins", "org.eclipse.core.runtime", ".settings", "org.eclipse.ui.workbench.prefs"),
    "savedTableConfigs.xml": os.path.join(".metadata", ".plugins", "com.heiler.ppm.std.ui", "savedTableConfigs.xml")
}

EXCLUDED_DIRS = {"Windows", "Program Files", "Program Files (x86)", "System Volume Information", "$RECYCLE.BIN"}

# Setup logging
os.makedirs(ARCHIVE_BASE, exist_ok=True)
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def find_installations(root_window):
    """Search C: for pim-desktop.exe, excluding system folders."""
    installations = []
    messagebox.showinfo("Search Starting", "Searching for Product 360 installations... This may take a few moments.", parent=root_window)
    for root, dirs, files in os.walk("C:\\", topdown=True):
        dirs[:] = [d for d in dirs if d not in EXCLUDED_DIRS]
        if "pim-desktop.exe" in files:
            installations.append(root)
        try:
            pass
        except PermissionError:
            logging.warning(f"Permission denied: {root}")
    return installations

def extract_workspace_dir(install_folder):
    """Read WORKSPACE_DIR from pim-desktop.cmd."""
    cmd_path = os.path.join(install_folder, "pim-desktop.cmd")
    if os.path.exists(cmd_path):
        with open(cmd_path, "r") as f:
            for line in f:
                if line.strip().startswith("SET WORKSPACE_DIR="):
                    return os.path.expandvars(line.split("=", 1)[1].strip())
    return None

def load_or_find_environments(root_window):
    """Load environments from config or find them, with user prompt if config exists."""
    if os.path.exists(CONFIG_FILE):
        if messagebox.askyesno("Use Saved Locations", "Use previously saved locations? (No = Search again)", parent=root_window):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        else:
            return search_and_save_environments(root_window)
    else:
        return search_and_save_environments(root_window)

def search_and_save_environments(root_window):
    """Search for installations and save to config with full paths."""
    environments = {}
    for install_folder in find_installations(root_window):
        workspace_dir = extract_workspace_dir(install_folder)
        if workspace_dir:
            environments[install_folder] = workspace_dir  # Full path as key (e.g., C:\PIM\Product 360 QA\client)
    with open(CONFIG_FILE, "w") as f:
        json.dump(environments, f, indent=4)
    return environments

def backup_files(env, environments, date_dropdown, selected_date, root_window):
    """Backup metadata files and refresh dates."""
    workspace_dir = environments[env]
    today = date.today().isoformat()
    backup_folder = os.path.join(ARCHIVE_BASE, env, today)
    os.makedirs(backup_folder, exist_ok=True)
    
    for filename, rel_path in SOURCE_FILES.items():
        source_path = os.path.join(workspace_dir, rel_path)
        dest_file = os.path.join(backup_folder, filename)
        if os.path.exists(source_path):
            shutil.copy2(source_path, dest_file)
            logging.info(f"Backed up {source_path} to {dest_file}")
        else:
            logging.warning(f"Source file not found: {source_path}")
    messagebox.showinfo("Backup Complete", f"Files backed up to {env}/{today}", parent=root_window)
    update_date_dropdown(env, date_dropdown, selected_date)

def clear_metadata(env, environments, root_window):
    """Delete the .metadata folder."""
    workspace_dir = environments[env]
    metadata_folder = os.path.join(workspace_dir, ".metadata")
    if not os.path.exists(metadata_folder):
        messagebox.showinfo("Clear Metadata", f"No .metadata folder found for {env}", parent=root_window)
        return
    
    if messagebox.askyesno("Confirm Clear", f"Delete .metadata folder for {env}? This cannot be undone.", parent=root_window):
        try:
            shutil.rmtree(metadata_folder)
            logging.info(f"Deleted {metadata_folder}")
            messagebox.showinfo("Clear Complete", f".metadata folder for {env} deleted", parent=root_window)
        except Exception as e:
            logging.error(f"Error deleting {metadata_folder}: {e}")
            messagebox.showerror("Clear Error", f"Error: {e}", parent=root_window)

def restore_files(env, date_str, environments, root_window):
    """Restore metadata files from a backup."""
    workspace_dir = environments[env]
    backup_folder = os.path.join(ARCHIVE_BASE, env, date_str)
    if not os.path.exists(backup_folder):
        messagebox.showerror("Restore Error", f"No backup found for {env} on {date_str}", parent=root_window)
        return
    
    overwrite = None
    for filename, rel_path in SOURCE_FILES.items():
        dest_path = os.path.join(workspace_dir, rel_path)
        if os.path.exists(dest_path):
            if overwrite is None:
                overwrite = messagebox.askyesno("Overwrite Files", "Some files exist. Overwrite them?", parent=root_window)
            if not overwrite:
                messagebox.showinfo("Restore Cancelled", "Restore operation cancelled.", parent=root_window)
                return
            break
    
    for filename, rel_path in SOURCE_FILES.items():
        source_file = os.path.join(backup_folder, filename)
        dest_path = os.path.join(workspace_dir, rel_path)
        if os.path.exists(source_file):
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            shutil.copy2(source_file, dest_path)
            logging.info(f"Restored {source_file} to {dest_path}")
        else:
            logging.warning(f"Backup file missing: {source_file}")
    messagebox.showinfo("Restore Complete", f"Files restored for {env} from {date_str}", parent=root_window)

def update_date_dropdown(env, date_dropdown, selected_date):
    """Update the backup date dropdown."""
    env_archive = os.path.join(ARCHIVE_BASE, env)
    if os.path.exists(env_archive):
        dates = [d for d in os.listdir(env_archive) if os.path.isdir(os.path.join(env_archive, d))]
        dates.sort(reverse=True)
        date_dropdown["values"] = dates
        selected_date.set(dates[0] if dates else "No backups found")
    else:
        date_dropdown["values"] = []
        selected_date.set("No backups found")

def add_manual_environment(environments, env_dropdown, selected_env, root_window):
    """Manually add an install location."""
    folder = askdirectory(title="Select Product 360 Install Folder", parent=root_window)
    if folder:
        workspace_dir = extract_workspace_dir(folder)
        if workspace_dir:
            if folder in environments:
                messagebox.showerror("Duplicate", f"{folder} already exists.", parent=root_window)
            else:
                environments[folder] = workspace_dir
                with open(CONFIG_FILE, "w") as f:
                    json.dump(environments, f, indent=4)
                env_dropdown["values"] = list(environments.keys())
                selected_env.set(folder)
                messagebox.showinfo("Environment Added", f"Added {folder}: {workspace_dir}", parent=root_window)
        else:
            messagebox.showerror("Invalid Folder", "No valid WORKSPACE_DIR found in pim-desktop.cmd", parent=root_window)

def create_gui(environments):
    """Create the GUI with a single Tk() instance."""
    root = Tk()
    root.title("Product 360 Layout Manager")
    root.geometry("400x200")

    # Full paths in dropdown (e.g., C:\PIM\Product 360 QA\client)
    env_list = list(environments.keys())

    ttk.Label(root, text="Select Environment:").pack(pady=5)
    selected_env = StringVar()
    env_dropdown = ttk.Combobox(root, textvariable=selected_env, values=env_list)
    env_dropdown.pack()
    if env_list:
        selected_env.set(env_list[0])

    ttk.Label(root, text="Select Backup Date:").pack(pady=5)
    selected_date = StringVar()
    date_dropdown = ttk.Combobox(root, textvariable=selected_date)
    date_dropdown.pack()
    if env_list:
        update_date_dropdown(selected_env.get(), date_dropdown, selected_date)

    def on_env_change(*args):
        update_date_dropdown(selected_env.get(), date_dropdown, selected_date)
    selected_env.trace("w", on_env_change)

    ttk.Button(root, text="Backup Now", command=lambda: backup_files(selected_env.get(), environments, date_dropdown, selected_date, root)).pack(pady=5)
    ttk.Button(root, text="Clear Metadata", command=lambda: clear_metadata(selected_env.get(), environments, root)).pack(pady=5)
    ttk.Button(root, text="Restore Metadata", command=lambda: restore_files(selected_env.get(), selected_date.get(), environments, root)).pack(pady=5)
    ttk.Button(root, text="Add Manual Location", command=lambda: add_manual_environment(environments, env_dropdown, selected_env, root)).pack(pady=5)

    root.mainloop()

def main():
    """Main entry point."""
    try:
        root = Tk()  # Create root for initial dialogs
        root.withdraw()  # Hide it until GUI is ready
        environments = load_or_find_environments(root)
        if not environments:
            messagebox.showerror("No Environments", "No Product 360 installations found. Please install or add manually.", parent=root)
            root.destroy()
            return
        root.destroy()  # Destroy temporary root
        create_gui(environments)
    except Exception as e:
        logging.error(f"Startup error: {traceback.format_exc()}")
        messagebox.showerror("Error", f"Failed to start: {e}")

if __name__ == "__main__":
    main()