import os
import pandas as pd
import glob
import re
from typing import List, Dict, Optional, Any

# Define the folder path
folder_path: str = r"C:\Reporting\Data Downloaded from IFS"

# Function to normalize column names for better matching
def normalize_column_name(name: str) -> str:
    """
    Normalizes a column name by converting to lowercase and removing whitespace.

    Args:
        name (str): The original column name.

    Returns:
        str: The normalized column name.
    """
    return re.sub(r'\s+', '', str(name)).lower()

# Function to find the best match for a column name
def find_column_match(df: pd.DataFrame, column_name: str) -> Optional[str]:
    """
    Finds the best matching column name in a DataFrame.
    Searches for exact normalized match first, then for partial matches.
    If multiple partial matches are found, a warning is issued, and the first one is returned.

    Args:
        df (pd.DataFrame): The DataFrame to search within.
        column_name (str): The target column name to find.

    Returns:
        Optional[str]: The actual column name from the DataFrame that matches, or None if no suitable match is found.
    """
    if df.empty:
        return None
        
    norm_target: str = normalize_column_name(column_name)
    norm_to_actual: Dict[str, str] = {normalize_column_name(col): col for col in df.columns}

    # Check for exact match after normalization
    if norm_target in norm_to_actual:
        return norm_to_actual[norm_target]

    # Check for partial matches
    possible_matches: List[str] = []
    for norm_col, actual_col in norm_to_actual.items():
        if norm_target in norm_col or norm_col in norm_target:
            possible_matches.append(actual_col)
    
    if possible_matches:
        if len(possible_matches) > 1:
            print(f"Warning: Multiple partial matches found for '{column_name}': {possible_matches}. Using '{possible_matches[0]}'.")
        return possible_matches[0]
    
    return None

# Function to read a file based on its extension
def read_file(file_path: str) -> pd.DataFrame:
    """
    Reads a data file (CSV, Excel, TXT, DAT) into a pandas DataFrame.
    Handles potential encoding issues for CSVs and attempts to infer separators for TXT/DAT files.

    Args:
        file_path (str): The path to the file.

    Returns:
        pd.DataFrame: The loaded DataFrame, or an empty DataFrame if reading fails.
    """
    _, ext = os.path.splitext(file_path)
    
    if ext.lower() in ['.csv']:
        try:
            return pd.read_csv(file_path)
        except UnicodeDecodeError:
            try:
                return pd.read_csv(file_path, encoding='latin1')
            except Exception as e:
                print(f"Error reading CSV {file_path} (even with latin1): {e}")
                return pd.DataFrame()
        except Exception as e:
            print(f"Error reading CSV {file_path}: {e}")
            return pd.DataFrame()
            
    elif ext.lower() in ['.xlsx', '.xls']:
        try:
            return pd.read_excel(file_path)
        except Exception as e:
            print(f"Error reading Excel file {file_path}: {e}")
            return pd.DataFrame()
            
    elif ext.lower() in ['.txt', '.dat']:
        # Try with different common separators.
        # This heuristic assumes that a valid delimited file will result in more than one column.
        for sep in ['\t', ',', ';', '|']:
            try:
                # Using engine='python' can be more robust for varied delimiters or bad lines
                df = pd.read_csv(file_path, sep=sep, engine='python') 
                if len(df.columns) > 1:
                    return df
            except Exception: # Continue to next separator if current one fails or file is malformed for this sep
                continue
        print(f"Could not determine separator or read {file_path} as a valid delimited text file with multiple columns.")
        return pd.DataFrame()
        
    else:
        print(f"Unsupported file format: {file_path}")
        return pd.DataFrame()

# Function to find AE, PT, and P files
def find_files(folder: str, file_type_keyword: str) -> List[str]:
    """
    Finds files in a folder that contain a specific keyword in their name and match given extensions.

    Args:
        folder (str): The directory to search in.
        file_type_keyword (str): The keyword to look for in filenames (e.g., "AE", "PT").

    Returns:
        List[str]: A list of paths to the found files.
    """
    all_files: List[str] = []
    extensions: List[str] = ['*.csv', '*.xlsx', '*.xls', '*.txt', '*.dat']
    for extension in extensions:
        search_pattern: str = os.path.join(folder, f"*{file_type_keyword}*{extension}")
        all_files.extend(glob.glob(search_pattern))
    return all_files

# Main data pull function
def pull_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    """
    Main function to pull and process data.
    - Finds and reads "AE", "PT", and "P" type files.
    - Concatenates files of the same type.
    - Extracts project manager mapping from "P" data.
    - Extracts and standardizes required columns from "AE" data, handling duplicates.

    Returns:
        tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]: 
            A tuple containing:
            - ae_data: Processed DataFrame from "AE" files.
            - pt_data: Combined DataFrame from "PT" files.
            - p_data: Combined DataFrame from "P" files.
            - project_manager_mapping: Dictionary mapping project IDs to manager descriptions.
    """
    # Find AE, PT, and P files
    ae_files: List[str] = find_files(folder_path, "AE")
    pt_files: List[str] = find_files(folder_path, "PT")
    p_files: List[str] = find_files(folder_path, "P")

    print(f"Found {len(ae_files)} AE files: {ae_files}")
    print(f"Found {len(pt_files)} PT files: {pt_files}")
    print(f"Found {len(p_files)} P files: {p_files}")

    # Read all AE files and concatenate them
    ae_data_frames: List[pd.DataFrame] = []
    for file in ae_files:
        print(f"Reading AE file: {file}")
        df = read_file(file)
        if not df.empty:
            print(f"  - Read {len(df)} rows with {len(df.columns)} columns from {file}")
            ae_data_frames.append(df)

    # Concatenate all AE data frames.
    # Note: pd.concat handles differing columns across DataFrames by filling missing values with NaN.
    if ae_data_frames:
        ae_data = pd.concat(ae_data_frames, ignore_index=True)
        print(f"Total AE records after concatenation: {len(ae_data)}")
    else:
        ae_data = pd.DataFrame()
        print("No AE data loaded.")

    # Read all PT files and concatenate them
    pt_data_frames: List[pd.DataFrame] = []
    for file in pt_files:
        print(f"Reading PT file: {file}")
        df = read_file(file)
        if not df.empty:
            print(f"  - Read {len(df)} rows with {len(df.columns)} columns from {file}")
            pt_data_frames.append(df)

    if pt_data_frames:
        pt_data = pd.concat(pt_data_frames, ignore_index=True)
        print(f"Total PT records after concatenation: {len(pt_data)}")
    else:
        pt_data = pd.DataFrame()
        print("No PT data loaded.")

    # Read all P files and concatenate them
    p_data_frames: List[pd.DataFrame] = []
    for file in p_files:
        print(f"Reading P file: {file}")
        df = read_file(file)
        if not df.empty:
            print(f"  - Read {len(df)} rows with {len(df.columns)} columns from {file}")
            p_data_frames.append(df)

    if p_data_frames:
        p_data = pd.concat(p_data_frames, ignore_index=True)
        print(f"Total P records after concatenation: {len(p_data)}")
    else:
        p_data = pd.DataFrame()
        print("No P data loaded.")

    # Extract project manager information from P data
    project_manager_mapping: Dict[str, str] = {}
    if not p_data.empty:
        try:
            # Define desired column names for P data
            project_col_target_name: str = 'Project'
            manager_desc_col_target_name: str = 'Manager Description'
            
            project_col_actual: Optional[str] = find_column_match(p_data, project_col_target_name)
            manager_desc_col_actual: Optional[str] = find_column_match(p_data, manager_desc_col_target_name)
            
            if project_col_actual and manager_desc_col_actual:
                print(f"Found project column '{project_col_actual}' (searched for '{project_col_target_name}') and manager description column '{manager_desc_col_actual}' (searched for '{manager_desc_col_target_name}') in P data.")
                
                for _, row in p_data.iterrows():
                    project_val = row[project_col_actual]
                    manager_desc_val = row[manager_desc_col_actual]
                    if pd.notna(project_val) and pd.notna(manager_desc_val):
                        project_manager_mapping[str(project_val).strip()] = str(manager_desc_val).strip()
                
                print(f"Created mapping for {len(project_manager_mapping)} projects to managers.")
            else:
                if not project_col_actual:
                    print(f"Warning: Could not find a column similar to '{project_col_target_name}' in P data.")
                if not manager_desc_col_actual:
                    print(f"Warning: Could not find a column similar to '{manager_desc_col_target_name}' in P data.")
        except Exception as e:
            print(f"Error processing P data for project manager mapping: {e}")
    else:
        print("No P data available to extract manager information.")

    # Extract required fields from AE data
    if not ae_data.empty:
        try:
            # Define required columns with standard names. The script will try to find matching actual column names.
            # These standard names will be the column names in the final ae_data DataFrame.
            ae_standard_to_actual_map: Dict[str, Optional[str]] = {
                'Activity Seq': None,
                'Project': None,
                'Project Description': None,
                'Activity': None,
                'Activity Description': None,
                'Estimated Revenue': None,
                'Estimated Cost': None
            }
            
            for standard_name in ae_standard_to_actual_map:
                actual_match: Optional[str] = find_column_match(ae_data, standard_name)
                if actual_match:
                    ae_standard_to_actual_map[standard_name] = actual_match
                    print(f"Matched standard column '{standard_name}'