import os
import shutil
import json
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Define source and destination directories
source_base = r'V:\mel_energy_projects\2 OPEN'
destination = r'C:\Reporting\Variations'
log_file = os.path.join(destination, "report_log.html")
variations_cache = os.path.join(destination, "variations_cache.json")
copied_files_cache = os.path.join(destination, "copied_files_cache.json")

# Ensure destination exists
os.makedirs(destination, exist_ok=True)

# Load existing variations structure
if os.path.exists(variations_cache):
    with open(variations_cache, "r") as cache_file:
        variations_map = json.load(cache_file)
else:
    variations_map = {}

# Load cache of previously copied files
if os.path.exists(copied_files_cache):
    with open(copied_files_cache, "r") as cache_file:
        copied_files = json.load(cache_file)
else:
    copied_files = {}

def find_variations_folder(source_folder):
    """ Find variations folder using the known folder structure pattern """
    if source_folder in variations_map:
        return variations_map[source_folder]
    
    # Known structure pattern:
    # Project folder -> Q Number folder -> 1 Management -> 6 Variations
    
    # First find 'Q' prefixed folders
    try:
        q_folders = []
        for item in os.listdir(source_folder):
            item_path = os.path.join(source_folder, item)
            if os.path.isdir(item_path) and (item.startswith('Q') or item.startswith('q')):
                q_folders.append(item_path)
            
        # If no Q folder found, try to find a folder that contains a Q folder
        if not q_folders:
            for item in os.listdir(source_folder):
                item_path = os.path.join(source_folder, item)
                if os.path.isdir(item_path):
                    try:
                        for sub_item in os.listdir(item_path):
                            sub_path = os.path.join(item_path, sub_item)
                            if os.path.isdir(sub_path) and (sub_item.startswith('Q') or sub_item.startswith('q')):
                                q_folders.append(sub_path)
                    except (PermissionError, FileNotFoundError):
                        continue
        
        # Now look for "1 Management" in each Q folder
        for q_folder in q_folders:
            mgmt_path = None
            try:
                for item in os.listdir(q_folder):
                    item_path = os.path.join(q_folder, item)
                    if os.path.isdir(item_path) and ("1 Management" in item or "01 Management" in item):
                        mgmt_path = item_path
                        break
            except (PermissionError, FileNotFoundError):
                continue
                
            if mgmt_path:
                # Look for "6 Variations" in Management folder
                try:
                    for item in os.listdir(mgmt_path):
                        item_path = os.path.join(mgmt_path, item)
                        if os.path.isdir(item_path) and ("6 Variations" in item or "06 Variations" in item):
                            print(f"Found Variations folder using pattern: {item_path}")
                            variations_map[source_folder] = item_path
                            with open(variations_cache, "w") as cache_file:
                                json.dump(variations_map, cache_file, indent=4)
                            return item_path
                except (PermissionError, FileNotFoundError):
                    continue
    
    except (PermissionError, FileNotFoundError) as e:
        print(f"Error accessing directory {source_folder}: {str(e)}")
        return None
    
    # Fallback to a simpler search for any "Variations" folder at levels 1 and 2
    try:
        # Try to find any Variations folder in immediate subdirectories
        for item in os.listdir(source_folder):
            item_path = os.path.join(source_folder, item)
            if not os.path.isdir(item_path):
                continue
                
            if "Variations" in item:
                print(f"Found Variations folder directly: {item_path}")
                variations_map[source_folder] = item_path
                with open(variations_cache, "w") as cache_file:
                    json.dump(variations_map, cache_file, indent=4)
                return item_path
                
            # Check one level deeper
            try:
                for sub_item in os.listdir(item_path):
                    sub_path = os.path.join(item_path, sub_item)
                    if not os.path.isdir(sub_path):
                        continue
                        
                    if "Variations" in sub_item:
                        print(f"Found Variations folder in subdirectory: {sub_path}")
                        variations_map[source_folder] = sub_path
                        with open(variations_cache, "w") as cache_file:
                            json.dump(variations_map, cache_file, indent=4)
                        return sub_path
            except (PermissionError, FileNotFoundError):
                continue
    except (PermissionError, FileNotFoundError):
        pass
    
    print(f"No Variations folder found for {source_folder}")
    return None

def copy_word_docs(source_folder):
    """ Copy Word documents from the Variations folder and its subfolders """
    variations_folder = find_variations_folder(source_folder)
    if not variations_folder:
        print(f'No Variations folder found for {source_folder}')
        return 0, 0
    
    print(f' --> Found "Variations" folder: {variations_folder}')
    file_count = 0
    files_skipped = 0
    
    # Create a key for this project in the copied_files dict if it doesn't exist
    project_name = os.path.basename(source_folder)
    if project_name not in copied_files:
        copied_files[project_name] = []
    
    # First try to find Word docs directly in the Variations folder
    for file in os.listdir(variations_folder):
        if file.endswith(('.doc', '.docx')):
            src_path = os.path.join(variations_folder, file)
            dest_path = os.path.join(destination, file)
            
            # Check if file was already copied before
            file_record = {
                "source": src_path,
                "destination": dest_path,
                "date_copied": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # Skip if this exact file was already copied
            if any(record["source"] == src_path for record in copied_files[project_name]):
                print(f' ----> Skipped (already copied): {src_path}')
                files_skipped += 1
                continue
                
            shutil.copy2(src_path, dest_path)
            copied_files[project_name].append(file_record)
            file_count += 1
            print(f' ----> Copied: {src_path} -> {dest_path}')
    
    # Then check subfolders (like "0 Q6941 Variation 1")
    for item in os.listdir(variations_folder):
        subfolder_path = os.path.join(variations_folder, item)
        if os.path.isdir(subfolder_path):
            print(f' --> Checking variation subfolder: {item}')
            try:
                for file in os.listdir(subfolder_path):
                    if file.endswith(('.doc', '.docx')):
                        src_path = os.path.join(subfolder_path, file)
                        
                        # Create a more descriptive filename to avoid conflicts
                        # Format: ProjectName_VariationNumber_OriginalFilename
                        variation_name = item.replace(" ", "_")
                        new_filename = f"{project_name}_{variation_name}_{file}"
                        
                        dest_path = os.path.join(destination, new_filename)
                        
                        # Check if file was already copied before
                        file_record = {
                            "source": src_path,
                            "destination": dest_path,
                            "date_copied": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        
                        # Skip if this exact file was already copied
                        if any(record["source"] == src_path for record in copied_files[project_name]):
                            print(f' ----> Skipped (already copied): {src_path}')
                            files_skipped += 1
                            continue
                            
                        shutil.copy2(src_path, dest_path)
                        copied_files[project_name].append(file_record)
                        file_count += 1
                        print(f' ----> Copied: {src_path} -> {dest_path}')
            except (PermissionError, FileNotFoundError) as e:
                print(f' ----> Error accessing subfolder {subfolder_path}: {str(e)}')
    
    # Save the updated copied files cache
    with open(copied_files_cache, "w") as cache_file:
        json.dump(copied_files, cache_file, indent=4)
    
    return 1, file_count, files_skipped

def generate_html_report(log_entries, total_folders_checked, total_files_copied, total_files_skipped=0):
    """ Generate HTML report """
    html_report = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Variations Copy Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 20px; }}
            .container {{ max-width: 800px; margin: auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); }}
            h2 {{ color: #333; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ padding: 10px; border: 1px solid #ddd; text-align: left; }}
            th {{ background-color: #007BFF; color: white; }}
            .summary {{ margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Variations Copy Report</h2>
            <p><strong>Date:</strong> {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
            
            <div class="summary">
                <h3>Summary</h3>
                <p><strong>Total Folders Checked:</strong> {total_folders_checked}</p>
                <p><strong>Total Files Copied:</strong> {total_files_copied}</p>
                <p><strong>Total Files Skipped (already copied):</strong> {total_files_skipped}</p>
            </div>
            
            <table>
                <tr>
                    <th>Project</th>
                    <th>Folders Checked</th>
                    <th>Files Copied</th>
                    <th>Files Skipped</th>
                </tr>
    '''
    for entry in log_entries:
        skipped = entry.get("files_skipped", 0)  # For backward compatibility
        html_report += f'<tr><td>{entry["project"]}</td><td>{entry["folders_checked"]}</td><td>{entry["files_copied"]}</td><td>{skipped}</td></tr>'
    html_report += f'''
            </table>
        </div>
    </body>
    </html>
    '''
    
    # Save the HTML report
    with open(log_file, "w") as f:
        f.write(html_report)
    
    print(f"Report generated at {log_file}")

class VariationsCopyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Variations Document Copy Tool")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        
        # Source and destination variables
        self.source_var = tk.StringVar(value=source_base)
        self.destination_var = tk.StringVar(value=destination)
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Source directory
        ttk.Label(settings_frame, text="Source Directory:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.source_var, width=50).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(settings_frame, text="Browse...", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)
        
        # Destination directory
        ttk.Label(settings_frame, text="Destination Directory:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.destination_var, width=50).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(settings_frame, text="Browse...", command=self.browse_destination).grid(row=1, column=2, padx=5, pady=5)
        
        # Load projects button
        ttk.Button(settings_frame, text="Load Projects", command=self.load_projects).grid(row=2, column=1, padx=5, pady=10)
        
        # Create projects frame
        projects_frame = ttk.LabelFrame(main_frame, text="Available Projects", padding="10")
        projects_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create treeview for projects
        self.tree = ttk.Treeview(projects_frame, columns=("path",), selectmode="extended")
        self.tree.heading("#0", text="Project Name")
        self.tree.heading("path", text="Path")
        self.tree.column("#0", width=200)
        self.tree.column("path", width=500)
        
        # Add scrollbar to treeview
        scrollbar_y = ttk.Scrollbar(projects_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(projects_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Layout with scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # Configure grid weights
        projects_frame.grid_rowconfigure(0, weight=1)
        projects_frame.grid_columnconfigure(0, weight=1)
        
        # Create action buttons frame
        actions_frame = ttk.Frame(main_frame, padding="10")
        actions_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Action buttons
        ttk.Button(actions_frame, text="Select All", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions_frame, text="Deselect All", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions_frame, text="Copy Selected Directly", command=self.copy_selected).pack(side=tk.RIGHT, padx=5)
        ttk.Button(actions_frame, text="Check Variations", command=self.check_variations).pack(side=tk.RIGHT, padx=5)
        ttk.Button(actions_frame, text="View Last Report", command=self.view_report).pack(side=tk.RIGHT, padx=5)
        
        # Add help button
        help_button = ttk.Button(actions_frame, text="?", width=3, command=self.show_help)
        help_button.pack(side=tk.LEFT, padx=15)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, padx=5, pady=5)
        
        # Auto-load projects after startup
        self.root.after(500, self.load_projects)
    
    def browse_source(self):
        directory = filedialog.askdirectory(initialdir=self.source_var.get())
        if directory:
            self.source_var.set(directory)
    
    def browse_destination(self):
        directory = filedialog.askdirectory(initialdir=self.destination_var.get())
        if directory:
            self.destination_var.set(directory)
            global destination
            destination = directory
            global log_file
            log_file = os.path.join(destination, "report_log.html")
    
    def load_projects(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        source_dir = self.source_var.get()
        if not os.path.isdir(source_dir):
            messagebox.showerror("Error", f"Source directory not found: {source_dir}")
            return
        
        self.status_var.set(f"Loading projects from {source_dir}...")
        self.root.update()
        
        # Get all project folders
        try:
            project_folders = []
            
            # Debug information
            try:
                dir_contents = os.listdir(source_dir)
                print(f"Directory contents of {source_dir}:")
                for item in dir_contents:
                    full_path = os.path.join(source_dir, item)
                    is_dir = os.path.isdir(full_path)
                    print(f" - {item} (Directory: {is_dir})")
                    if is_dir:
                        project_folders.append(full_path)
            except Exception as e:
                print(f"Error listing directory: {str(e)}")
                messagebox.showerror("Error", f"Failed to list directory contents: {str(e)}")
                self.status_var.set("Error loading projects")
                return
            
            if not project_folders:
                messagebox.showinfo("Information", "No project folders found in the source directory.")
                self.status_var.set("No projects found")
                return
            
            for idx, project in enumerate(project_folders):
                project_name = os.path.basename(project)
                self.tree.insert("", tk.END, text=project_name, values=(project,))
            
            self.status_var.set(f"Loaded {len(project_folders)} projects")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load projects: {str(e)}")
            self.status_var.set("Error loading projects")
    
    def select_all(self):
        for item in self.tree.get_children():
            self.tree.selection_add(item)
    
    def deselect_all(self):
        for item in self.tree.get_children():
            self.tree.selection_remove(item)
    
    def check_variations(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Information", "No projects selected")
            return
        
        self.status_var.set("Checking for Variations folders...")
        self.root.update()
        
        found_count = 0
        projects_checked = 0
        
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Checking Variations Folders")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Progress info
        ttk.Label(progress_window, text="Searching for Variations folders...").pack(pady=10)
        progress_var = tk.StringVar(value="")
        ttk.Label(progress_window, textvariable=progress_var).pack(pady=5)
        
        # Progress bar
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10, padx=20, fill=tk.X)
        progress["maximum"] = len(selected_items)
        
        # Cancel button
        self.cancel_check = False
        cancel_button = ttk.Button(progress_window, text="Cancel", command=lambda: setattr(self, 'cancel_check', True))
        cancel_button.pack(pady=10)
        
        # Process each project
        for i, item in enumerate(selected_items):
            if self.cancel_check:
                break
                
            project_path = self.tree.item(item, "values")[0]
            project_name = self.tree.item(item, "text")
            projects_checked += 1
            
            progress_var.set(f"Checking: {project_name}")
            progress["value"] = i
            progress_window.update()
            
            variations_folder = find_variations_folder(project_path)
            if variations_folder:
                self.tree.item(item, tags=("found",))
                found_count += 1
            else:
                self.tree.item(item, tags=("notfound",))
        
        # Close progress window
        progress_window.destroy()
        
        self.tree.tag_configure("found", background="#90EE90")  # Light green
        self.tree.tag_configure("notfound", background="#FFB6C1")  # Light red
        
        self.status_var.set(f"Check complete. Found {found_count} Variations folders out of {projects_checked} projects.")
        
        # Show results message
        messagebox.showinfo("Check Complete", 
                            f"Found {found_count} Variations folders out of {projects_checked} projects.\n\n"
                            f"Green: Has Variations folder\n"
                            f"Red: No Variations folder found")
    
    def copy_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Information", "No projects selected")
            return
        
        dest_dir = self.destination_var.get()
        if not os.path.isdir(dest_dir):
            try:
                os.makedirs(dest_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create destination directory: {str(e)}")
                return
        
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Copying Variation Documents")
        progress_window.geometry("400x180")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Progress info
        ttk.Label(progress_window, text="Copying documents...").pack(pady=10)
        progress_var = tk.StringVar(value="")
        progress_label = ttk.Label(progress_window, textvariable=progress_var)
        progress_label.pack(pady=5)
        
        # Add a label for skipped files
        skipped_var = tk.StringVar(value="")
        skipped_label = ttk.Label(progress_window, textvariable=skipped_var)
        skipped_label.pack(pady=5)
        
        # Progress bar
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10, padx=20, fill=tk.X)
        progress["maximum"] = len(selected_items)
        
        # Cancel button
        self.cancel_copy = False
        cancel_button = ttk.Button(progress_window, text="Cancel", command=lambda: setattr(self, 'cancel_copy', True))
        cancel_button.pack(pady=10)
        
        self.status_var.set("Copying documents...")
        self.root.update()
        
        total_folders_checked = 0
        total_files_copied = 0
        total_files_skipped = 0
        log_entries = []
        
        for i, item in enumerate(selected_items):
            if self.cancel_copy:
                break
                
            project_path = self.tree.item(item, "values")[0]
            project_name = self.tree.item(item, "text")
            
            progress_var.set(f"Processing: {project_name}")
            skipped_var.set(f"Files skipped: {total_files_skipped}")
            progress["value"] = i
            progress_window.update()
            
            # Mark as in progress
            self.tree.item(item, tags=("inprogress",))
            self.tree.tag_configure("inprogress", background="#FFFFCC")  # Light yellow
            progress_window.update()
            
            folders_checked, files_copied, files_skipped = copy_word_docs(project_path)
            total_folders_checked += folders_checked
            total_files_copied += files_copied
            total_files_skipped += files_skipped
            
            # Mark result
            if files_copied > 0:
                self.tree.item(item, tags=("copied",))
            elif files_skipped > 0 and files_copied == 0:
                self.tree.item(item, tags=("skipped",))
            elif folders_checked > 0:
                self.tree.item(item, tags=("empty",))
            else:
                self.tree.item(item, tags=("notfound",))
            
            log_entries.append({
                "project": project_path,
                "folders_checked": folders_checked,
                "files_copied": files_copied,
                "files_skipped": files_skipped
            })
        
        # Close progress window
        progress_window.destroy()
        
        # Configure tags
        self.tree.tag_configure("copied", background="#90EE90")  # Light green
        self.tree.tag_configure("skipped", background="#ADD8E6")  # Light blue
        self.tree.tag_configure("empty", background="#FFFFE0")  # Light yellow
        self.tree.tag_configure("notfound", background="#FFB6C1")  # Light red
        
        # Generate report
        generate_html_report(log_entries, total_folders_checked, total_files_copied, total_files_skipped)
        
        self.status_var.set(f"Copying complete. Copied {total_files_copied} files, skipped {total_files_skipped} files.")
        
        result_message = (
            f"Successfully copied {total_files_copied} files from {total_folders_checked} folders.\n"
            f"Skipped {total_files_skipped} previously copied files.\n\n"
            f"Green: Files copied\n"
            f"Blue: All files were already copied (skipped)\n"
            f"Yellow: Variation folder found but no Word documents\n"
            f"Red: No Variations folder found\n\n"
            f"Report saved to {log_file}"
        )
        messagebox.showinfo("Complete", result_message)
    
    def view_report(self):
        if os.path.exists(log_file):
            os.startfile(log_file)
        else:
            messagebox.showinfo("Information", "No report file found. Run a copy operation first.")
            
    def show_help(self):
        help_text = """
Variations Document Copy Tool - Help

This tool helps you copy Word documents from Variations folders in multiple projects.

Workflow:
1. Select projects from the list (use Ctrl+click for multiple selections)
2. Choose one of these options:

   • "Check Variations": 
     - Identifies which projects have Variations folders
     - Color codes the projects (green = has Variations folder, red = no Variations folder)
     - This is optional but helps you see which projects to focus on

   • "Copy Selected Directly": 
     - Immediately searches for and copies Word documents from Variations folders
     - Color codes after copying (green = files copied, yellow = empty Variations folder, red = no Variations folder)
     - Generates a report

"Copy Selected Directly" is the faster option if you know which projects to copy from.

Color Codes:
• Green: Has Variations folder / Files copied
• Yellow: Variations folder found but empty
• Red: No Variations folder found

Tips:
• The search now follows your exact folder structure pattern:
  Project > Q Number Folder > 1 Management > 6 Variations
• This targeted search is much faster than scanning all folders
• Previously found Variations folders are cached for faster future operations
        """
        help_window = tk.Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("600x500")
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        
        close_button = ttk.Button(help_window, text="Close", command=help_window.destroy)
        close_button.pack(pady=10)

def main():
    root = tk.Tk()
    app = VariationsCopyApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()