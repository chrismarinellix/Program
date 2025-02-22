import pandas as pd
import os
import openpyxl

# Define file paths
input_file_ae = r'C:\Users\chris\Downloads\ae.xlsx'
output_file_ae = r'C:\py\Program\Data\ae.xlsx'  # Save to the specified folder with the same name

input_file_p = r'C:\Users\chris\Downloads\P.xlsx'
output_file_p = r'C:\py\Program\Data\P.xlsx'  # Save to the specified folder with the same name

input_file_pt = r'C:\Users\chris\Downloads\PT.xlsx'
output_file_pt = r'C:\py\Program\Data\PT.xlsx'  # Save to the specified folder with the same name

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

    # Read the P Excel file
    df_p = pd.read_excel(input_file_p)
    
    # Rename the header from 'Project ID' to 'Project'
    df_p = df_p.rename(columns={'Project ID': 'Project'})
    
    # Save the modified P data to the specified folder with the same name
    df_p.to_excel(output_file_p, index=False)
    print(f"Modified P data has been saved to {output_file_p}")

    # Move the PT Excel file to the specified folder without making any changes
    if os.path.exists(input_file_pt):
        os.replace(input_file_pt, output_file_pt)
        print(f"PT file has been moved to {output_file_pt}")
    else:
        print(f"Error: The file {input_file_pt} was not found.")

except FileNotFoundError as e:
    print(f"Error: {e.filename} was not found.")
except Exception as e:
    print(f"An error occurred: {str(e)}")