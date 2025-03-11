import pandas as pd
from pathlib import Path

# Define file paths
base_path = Path(r"C:\Reporting\Py")
input_files = {
    "AE": base_path / "AE.xlsx",
    "P": base_path / "P.xlsx",
    "PT": base_path / "PT.xlsx"
}
output_file = base_path / "Project_Report_By_Manager.xlsx"

# Define the output headers
output_headers = [
    "Project", "Project Description", "Sub Project", "Sub Project Description",
    "Activity", "Activity Description", "Activity Status", "Cost/Revenue Element",
    "Estimated Cost", "Estimated Revenue", "Estimated Hours", "Cost To Complete",
    "Created By", "Modified By"
]

# Function to read and process each spreadsheet
def process_spreadsheet(file_path):
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None

    # Rename columns to standardize them (based on AE example)
    rename_dict = {
        "PROJECT": "Project",
        "PROJECT_DESCRIPTION": "Project Description",
        "SUB_PROJECT": "Sub Project",
        "SUB_PROJECT_DESCRIPTION": "Sub Project Description",
        "ACTIVITY": "Activity",
        "ACTIVITY_DESCRIPTION": "Activity Description",
        "ACTIVITY_STATUS": "Activity Status",
        "COST/REVENUE_ELEMENT": "Cost/Revenue Element",
        "ESTIMATED_COST": "Estimated Cost",
        "ESTIMATED_REVENUE": "Estimated Revenue",
        "ESTIMATED_HOURS": "Estimated Hours",
        "COST_TO_COMPLETE": "Cost To Complete",
        "CREATED_BY": "Created By",
        "MODIFIED_BY": "Modified By",
        "ACTIVITY_SEQ": "Activity Seq",
        # Assuming project manager column exists; adjust name if different
        "TUN_SHAMSUDDIN": "Project Manager"  # Replace with actual column name if different
    }

    df = df.rename(columns=rename_dict)

    # Keep only the columns we need (if they exist)
    columns_to_keep = list(rename_dict.values())
    df = df[[col for col in columns_to_keep if col in df.columns]]

    return df

# Process all spreadsheets and combine them
all_data = []
for sheet_name, file_path in input_files.items():
    df = process_spreadsheet(file_path)
    if df is not None:
        # Add a column to track the source file (optional, for debugging)
        df["Source"] = sheet_name
        all_data.append(df)

# Combine all data into a single DataFrame
if not all_data:
    print("No data loaded from any spreadsheet. Exiting.")
    exit()

combined_df = pd.concat(all_data, ignore_index=True)

# Handle the Estimated Cost and Estimated Revenue (they're in separate rows)
# Group by Activity Seq and pivot the Cost/Revenue Element
pivot_df = combined_df.pivot_table(
    index=["Project", "Project Description", "Sub Project", "Sub Project Description",
           "Activity", "Activity Description", "Activity Status", "Activity Seq",
           "Created By", "Modified By", "Project Manager"],
    columns="Cost/Revenue Element",
    values=["Estimated Cost", "Estimated Revenue", "Estimated Hours", "Cost To Complete"],
    aggfunc="first"
).reset_index()

# Flatten the multi-index columns
pivot_df.columns = [
    f"{col[0]} {col[1]}".strip() if col[1] else col[0]
    for col in pivot_df.columns
]

# Rename the pivoted columns to match output headers
pivot_df = pivot_df.rename(columns={
    "Estimated Cost COST": "Estimated Cost",
    "Estimated Revenue REVENUE": "Estimated Revenue",
    "Estimated Hours COST": "Estimated Hours",  # Adjust based on actual data
    "Cost To Complete COST": "Cost To Complete"
})

# Keep only the columns we need for the output
final_columns = [col for col in output_headers if col in pivot_df.columns]
final_df = pivot_df[final_columns]

# Add missing columns with NaN if they don't exist
for col in output_headers:
    if col not in final_df.columns:
        final_df[col] = pd.NA

# Reorder columns to match output_headers
final_df = final_df[output_headers]

# Group by Project Manager and sort projects alphabetically within each group
grouped = final_df.groupby("Project Manager")

# Create a new Excel workbook with a tab for each project manager
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for manager, group in grouped:
        # Sort by Project alphabetically
        group_sorted = group.sort_values(by="Project")
        
        # Drop the Project Manager column since it's the tab name
        group_sorted = group_sorted.drop(columns=["Project Manager"], errors="ignore")
        
        # Write to a new sheet named after the project manager
        # Replace invalid characters in sheet name
        sheet_name = str(manager).replace("/", "_").replace("\\", "_")[:31]  # Excel sheet name limit
        group_sorted.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Report generated successfully at {output_file}")