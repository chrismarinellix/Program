import pandas as pd
import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
import json

CONFIG_FILE = 'config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def consolidate_ae(input_file_ae, output_file_ae):
    try:
        # Read the AE Excel file
        df_ae = pd.read_excel(input_file_ae)
        
        # Iterate through the DataFrame and consolidate rows based on Activity Seq
        for activity_seq in df_ae['Activity Seq'].unique():
            # Get all rows with the same Activity Seq
            rows = df_ae[df_ae['Activity Seq'] == activity_seq]
            
            # If there are multiple rows, consolidate them
            if len(rows) > 1:
                # Take the first row as the base
                base_row_index = rows.index[0]
                base_row = rows.iloc[0]
                
                # Iterate through the remaining rows and consolidate data
                for _, row in rows.iloc[1:].iterrows():
                    if pd.isna(df_ae.at[base_row_index, 'Estimated Revenue']) and not pd.isna(row['Estimated Revenue']):
                        df_ae.at[base_row_index, 'Estimated Revenue'] = row['Estimated Revenue']
                    if pd.isna(df_ae.at[base_row_index, 'Estimated Cost']) and not pd.isna(row['Estimated Cost']):
                        df_ae.at[base_row_index, 'Estimated Cost'] = row['Estimated Cost']
                    if pd.isna(df_ae.at[base_row_index, 'Estimated Hours']) and not pd.isna(row['Estimated Hours']):
                        df_ae.at[base_row_index, 'Estimated Hours'] = row['Estimated Hours']
                    if pd.isna(df_ae.at[base_row_index, 'Estimated Cost To Complete']) and not pd.isna(row['Estimated Cost To Complete']):
                        df_ae.at[base_row_index, 'Estimated Cost To Complete'] = row['Estimated Cost To Complete']
                
                # Drop the duplicate rows except the first one
                df_ae = df_ae.drop(rows.index[1:])
        
        # Save the consolidated AE data to the specified folder with the same name
        try:
            # Attempt to close the file if it is open
            if os.path.exists(output_file_ae):
                wb = openpyxl.load_workbook(output_file_ae)
                wb.close()
            
            df_ae.to_excel(output_file_ae, index=False)
            print(f"Consolidated data has been saved to {output_file_ae}")
        except PermissionError:
            # If the file is open, save to a temporary file and then replace the original file
            temp_output_file_ae = output_file_ae.replace('.xlsx', '_temp.xlsx')
            df_ae.to_excel(temp_output_file_ae, index=False)
            os.replace(temp_output_file_ae, output_file_ae)
            print(f"Consolidated data has been saved to {output_file_ae} (file was open, so it was replaced)")
    except FileNotFoundError:
        print(f"Error: The file {input_file_ae} was not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def modify_p(input_file_p, output_file_p):
    try:
        # Read the P Excel file
        df_p = pd.read_excel(input_file_p)
        
        # Rename the header from 'Project ID' to 'Project'
        df_p = df_p.rename(columns={'Project ID': 'Project'})
        
        # Save the modified P data to the specified folder with the same name
        df_p.to_excel(output_file_p, index=False)
        print(f"Modified P data has been saved to {output_file_p}")
    except FileNotFoundError:
        print(f"Error: The file {input_file_p} was not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def move_pt(input_file_pt, output_file_pt):
    try:
        # Move the PT Excel file to the specified folder without making any changes
        if os.path.exists(input_file_pt):
            os.replace(input_file_pt, output_file_pt)
            print(f"PT file has been moved to {output_file_pt}")
        else:
            print(f"Error: The file {input_file_pt} was not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def execute_script():
    input_file_ae = ae_input_path.get()
    output_file_ae = ae_output_path.get()
    input_file_p = p_input_path.get()
    output_file_p = p_output_path.get()
    input_file_pt = pt_input_path.get()
    output_file_pt = pt_output_path.get()
    
    # Save the paths to the config file
    config = {
        "ae_input_path": input_file_ae,
        "ae_output_path": output_file_ae,
        "p_input_path": input_file_p,
        "p_output_path": output_file_p,
        "pt_input_path": input_file_pt,
        "pt_output_path": output_file_pt
    }
    save_config(config)
    
    consolidate_ae(input_file_ae, output_file_ae)
    modify_p(input_file_p, output_file_p)
    move_pt(input_file_pt, output_file_pt)
    messagebox.showinfo("Success", "Script executed successfully!")

def browse_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

# Load the last used paths from the config file
config = load_config()

# Create the main window
root = tk.Tk()
root.title("File Path Configuration")

# AE file paths
tk.Label(root, text="AE Input File Path:").grid(row=0, column=0, sticky=tk.W)
ae_input_path = tk.Entry(root, width=50)
ae_input_path.grid(row=0, column=1)
ae_input_path.insert(0, config.get("ae_input_path", r'C:\Users\chris\Downloads\ae.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_file(ae_input_path)).grid(row=0, column=2)

tk.Label(root, text="AE Output File Path:").grid(row=1, column=0, sticky=tk.W)
ae_output_path = tk.Entry(root, width=50)
ae_output_path.grid(row=1, column=1)
ae_output_path.insert(0, config.get("ae_output_path", r'C:\py\Program\Data\ae.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_folder(ae_output_path)).grid(row=1, column=2)

# P file paths
tk.Label(root, text="P Input File Path:").grid(row=2, column=0, sticky=tk.W)
p_input_path = tk.Entry(root, width=50)
p_input_path.grid(row=2, column=1)
p_input_path.insert(0, config.get("p_input_path", r'C:\Users\chris\Downloads\P.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_file(p_input_path)).grid(row=2, column=2)

tk.Label(root, text="P Output File Path:").grid(row=3, column=0, sticky=tk.W)
p_output_path = tk.Entry(root, width=50)
p_output_path.grid(row=3, column=1)
p_output_path.insert(0, config.get("p_output_path", r'C:\py\Program\Data\P.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_folder(p_output_path)).grid(row=3, column=2)

# PT file paths
tk.Label(root, text="PT Input File Path:").grid(row=4, column=0, sticky=tk.W)
pt_input_path = tk.Entry(root, width=50)
pt_input_path.grid(row=4, column=1)
pt_input_path.insert(0, config.get("pt_input_path", r'C:\Users\chris\Downloads\PT.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_file(pt_input_path)).grid(row=4, column=2)

tk.Label(root, text="PT Output File Path:").grid(row=5, column=0, sticky=tk.W)
pt_output_path = tk.Entry(root, width=50)
pt_output_path.grid(row=5, column=1)
pt_output_path.insert(0, config.get("pt_output_path", r'C:\py\Program\Data\PT.xlsx'))
tk.Button(root, text="Browse", command=lambda: browse_folder(pt_output_path)).grid(row=5, column=2)

# Execute button
tk.Button(root, text="Execute Script", command=execute_script).grid(row=6, column=1, pady=10)

# Run the main loop
root.mainloop()