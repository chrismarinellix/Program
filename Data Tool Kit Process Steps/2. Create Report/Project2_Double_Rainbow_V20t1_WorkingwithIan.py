import pandas as pd
import numpy as np
import warnings
import pickle
import os
from typing import Dict, Optional, List, Any
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime
import json
import sys

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('project_analyzer.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Suppress specific pandas warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style")

class ProjectAnalyzer:
    def __init__(self, file_paths: Dict[str, str], cache_file: str = 'data_cache.pkl'):
        self.file_paths = file_paths
        self.cache_file = cache_file
        self.dataframes: Dict[str, pd.DataFrame] = {}
        
        self.header_mapping = {
            'Project': [
                'Project',
                'Project ID'
            ],
            'Project Description': [
                'Project Description'
            ],
            'Manager Description': [
                'Manager Description'
            ],
            'Activity Status': [
                'Activity Status'
            ],
            'Status': [
                'Status'
            ],
            'Sales Quantity': [
                'Sales Quantity'
            ],
            'Invoice Quantity': [
                'Invoice Quantity'
            ],
            'Internal Quantity': [
                'Internal Quantity'
            ],
            'Invoice Status': [
                'Invoice Status'
            ],
            'Activity Sequence': [
                'Activity Sequence',
                'Activity Seq'
            ],
            'Activity': [
                'Activity',
                'Activity ID'
            ],
            'Activity Description': [
                'Activity Description',
                'Activity Desc',
                'Task Description'
            ],
            'Estimated Revenue': [
                'Estimated Revenue',
                'Budget'
            ],
            'Actual Revenue': [
                'Posted Revenue',
                'Sales Amount'
            ],
            'Sales Amount': [
                'Sales Amount',
                'Amount',
                'Total Sales'
            ]
        }

    def standardize_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        """Standardize DataFrame headers using mapping."""
        # Create reverse mapping with error handling
        logger.info(f"Original headers: {df.columns.tolist()}")
        print(f"Standardizing headers for DataFrame with columns: {df.columns.tolist()}")
        
        reverse_mapping = {}
        for standard, variants in self.header_mapping.items():
            for variant in variants:
                if variant in df.columns:  # Only map existing columns
                    reverse_mapping[variant] = standard
                    logger.info(f"Mapped {variant} to {standard}")
                    print(f"Found and mapped: {variant} -> {standard}")
        
        print(f"\nFinal mapping to be applied: {reverse_mapping}")
        
        # Log the mapping process
        logger.debug(f"Original columns: {df.columns.tolist()}")
        df = df.rename(columns=reverse_mapping)
        logger.debug(f"Mapped columns: {df.columns.tolist()}")
        print(f"Columns after mapping: {df.columns.tolist()}")
        
        if 'Activity' in df.columns:
            logger.info(f"Activity column found after mapping with values: {df['Activity'].unique()}")
        
        return df

    def load_data(self, force_refresh: bool = False) -> bool:
        """Load data from cache or files."""
        if os.path.exists(self.cache_file) and not force_refresh:
            cache_time = datetime.fromtimestamp(os.path.getmtime(self.cache_file))
            print(f"Found cache file from: {cache_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            try:
                with open(self.cache_file, 'rb') as f:
                    self.dataframes = pickle.load(f)
                logger.info("Data loaded from cache successfully")
                print(f"Loaded {len(self.dataframes)} dataframes from cache")
                for name in self.dataframes.keys():
                    print(f"- {name}.xlsx")
                return True
            except Exception as e:
                logger.error(f"Error loading cache: {e}")
                print("Cache file exists but couldn't be loaded, refreshing data...")
        else:
            if force_refresh:
                print("Forced refresh requested, loading fresh data...")
            else:
                print("No cache file found, loading fresh data...")
        
        return self._load_fresh_data()

    def _load_fresh_data(self) -> bool:
        """Load fresh data from Excel files."""
        try:
            # Load files directly
            for name, path in self.file_paths.items():
                logger.info(f"Loading {name} from {path}")
                df = pd.read_excel(path)
                # Standardize headers
                df = self.standardize_headers(df)
                self.dataframes[name] = df
            
            self._save_cache()
            return True
        except Exception as e:
            logger.error(f"Error loading fresh data: {e}")
            return False

    def _save_cache(self):
        """Save dataframes to cache file."""
        try:
            with open(self.cache_file, 'wb') as f:
                pickle.dump(self.dataframes, f)
            logger.info("Data cached successfully")
        except Exception as e:
            logger.error(f"Error saving cache: {e}")

    def _ensure_file_closed(self, file_path: str) -> bool:
        """Ensure specific Excel file is closed."""
        try:
            import win32com.client
            
            abs_path = os.path.abspath(file_path)
            excel = win32com.client.GetObject(Class="Excel.Application")
            
            for wb in excel.Workbooks:
                if os.path.abspath(wb.FullName) == abs_path:
                    print(f"Found open workbook: {wb.Name}")
                    wb.Close(SaveChanges=False)
                    print(f"Closed workbook: {wb.Name}")
                    
            if excel.Workbooks.Count == 0:
                excel.Quit()
                
            return True
            
        except Exception as e:
            logger.debug(f"Excel interaction: {str(e)}")
            return True

    def generate_weekly_report(self, output_path: str) -> bool:
        """Generate and save weekly report with dated copy."""
        try:
            if 'P' not in self.dataframes:
                print("Error: P file not found in loaded data")
                return False
            
            # Setup directories
            base_dir = os.path.dirname(output_path)
            archive_dir = os.path.join(base_dir, "Weekly Report")
            os.makedirs(archive_dir, exist_ok=True)
            logger.info("Directory setup completed successfully")
            
            # Generate timestamp for archive copy
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            archive_path = os.path.join(archive_dir, f"Weekly_Report_{timestamp}.xlsx")
            
            if not self._ensure_file_closed(output_path):
                print(f"Could not access {output_path} - please close Excel and try again")
                return False
            
            p_file = self.dataframes['P']
            
            # Debug: Print available columns
            logger.info(f"Available columns in P file: {p_file.columns.tolist()}")
            
            # Check if required columns exist
            required_columns = ['Project', 'Manager Description']
            missing_columns = [col for col in required_columns if col not in p_file.columns]
            
            if missing_columns:
                error_msg = f"Required columns missing from P file: {missing_columns}"
                logger.error(error_msg)
                print(error_msg)
                return False
                
            project_manager_map = p_file[required_columns].drop_duplicates()
            # After:
            managers = sorted([m for m in p_file['Manager Description'].dropna().unique() 
                            if m not in ['Grant', 'Ian', 'Gemma']])
            logger.info(f"Found {len(managers)} managers to process")
            
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                for manager in managers:
                    manager_projects = project_manager_map[
                        project_manager_map['Manager Description'] == manager
                    ]['Project'].unique()
                    
                    all_manager_data = []
                    for df_name, df in self.dataframes.items():
                        if 'Project' not in df.columns:
                            continue
                        
                        df_filtered = df[df['Project'].isin(manager_projects)].copy()
                        if not df_filtered.empty:
                            if 'Status' in df_filtered.columns:
                                df_filtered = df_filtered[
                                    df_filtered['Status'].astype(str).str.lower() == 'started'
                                ]
                            if not df_filtered.empty:
                                logger.info(f"Adding {len(df_filtered)} rows from {df_name} for {manager}")
                                all_manager_data.append(df_filtered)
                    
                    if all_manager_data:
                        manager_report = self._process_manager_data(manager, all_manager_data)
                        if len(manager_report) > 0:
                            self._write_manager_sheet(manager, manager_report, writer.book, writer)
                
                print(f"Generated report with tabs for {len(managers)} managers")
            
            # Create archive copy
            try:
                import shutil
                shutil.copy2(output_path, archive_path)
                print(f"Archive copy created successfully")
            except Exception as e:
                logger.error(f"Failed to create archive copy: {str(e)}")
                
            return True
                
        except Exception as e:
            print(f"Error generating weekly report: {str(e)}")
            logger.error(f"Error encountered: {str(e)}", exc_info=True)
            return False

    def inspect_data(self) -> None:
        """Display summary of data from each Excel file."""
        for name, df in self.dataframes.items():
            print(f"\n{'='*50}")
            print(f"File: {name}.xlsx")
            print(f"{'='*50}")
            print(f"Total rows: {len(df)}")
            print("\nColumns:")
            
            for col in df.columns:
                try:
                    non_null = df[col].count()
                    if non_null > 0:
                        sample = str(df[col].iloc[0])
                        if pd.isna(df[col].iloc[0]):
                            sample = 'nan'
                    else:
                        sample = 'N/A'
                    print(f"- {col}: {non_null} non-null values, Sample: {sample}")
                except Exception as e:
                    print(f"- {col}: Error getting sample - {str(e)}")
            
            if 'Manager Description' in df.columns:
                print("\nUnique Manager Descriptions:")
                try:
                    unique_managers = df['Manager Description'].dropna().unique()
                    for manager in unique_managers[:10]:  # Show first 10 managers
                        print(f"- {manager}")
                    if len(unique_managers) > 10:
                        print(f"... and {len(unique_managers) - 10} more")
                except Exception as e:
                    print(f"Error listing managers: {str(e)}")

    def debug_df_info(self, df_name: str):
        """Print debug information about a dataframe."""
        if df_name not in self.dataframes:
            print(f"DataFrame {df_name} not found!")
            return
            
        df = self.dataframes[df_name]
        print(f"\nDebug info for {df_name}:")
        print(f"Shape: {df.shape}")
        print("\nColumns:")
        for i, col in enumerate(df.columns):
            count = df[col].count()
            sample = df[col].iloc(0) if count > 0 else "No data"
            print(f"{i}: {col} ({count} non-null values) Sample: {sample}")

    def _resolve_duplicate_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Resolves duplicate columns by appending _N to duplicates."""
        cols = df.columns.tolist()
        new_cols = []
        seen = {}
        for col in cols:
            if col in seen:
                seen[col] += 1
                new_col = f"{col}_{seen[col]}"
                new_cols.append(new_col)
            else:
                seen[col] = 0
                new_cols.append(col)
        df.columns = new_cols
        return df

    def _safe_concat_dataframes(self, dataframes: List[pd.DataFrame]) -> pd.DataFrame:
        """Safely concatenate DataFrames with proper column handling."""
        try:
            # Get all columns at once
            all_columns = pd.Index(sorted(set().union(*(df.columns for df in dataframes))))
            
            # Normalize all dataframes at once
            normalized_dfs = []
            for df in dataframes:
                missing_cols = all_columns.difference(df.columns)
                if not missing_cols.empty:
                    df = df.reindex(columns=all_columns, fill_value=pd.NA)
                normalized_dfs.append(df)
            
            # Concatenate
            result = pd.concat(normalized_dfs, ignore_index=True)
            return result.fillna(pd.NA)
            
        except Exception as e:
            self.logger.error(f"Error concatenating dataframes: {str(e)}", exc_info=True)
            raise

    def _write_manager_sheet(self, manager: str, manager_report: pd.DataFrame, workbook: object, writer: object) -> None:
        """Write manager data to Excel sheet with enhanced column handling."""
        try:
            # Sheet setup
            first_name = str(manager).split()[0]
            sheet_name = first_name[:31].replace('/', '_').replace('\\', '_')
            logger.info(f"Creating sheet for {manager} with {len(manager_report)} rows")

            # Define columns upfront
            columns = [
                'Project',
                'Project Description',
                'Activity Sequence',
                'Activity',
                'Activity Description',
                'Status',
                'Billing Type',
                'Estimated Revenue',
                'Actual Revenue',
                'Remaining Budget'
            ]

            # Create new data structure with expanded activities
            expanded_data = []
            current_project = None
            
            # Sort by Project Description
            manager_report = manager_report.sort_values(['Project Description']).fillna('')
            
            # Define terms that indicate Fixed billing
            fixed_terms = ['fixed', 'hrs', 'hours']
            
            # Process each row and expand activities
            for _, row in manager_report.iterrows():
                project_desc = row['Project Description']
                
                # Add header rows for new project
                if project_desc != current_project:
                    if current_project is not None:
                        # Add three blank rows between projects
                        for _ in range(3):
                            expanded_data.append({col: '' for col in columns})
                    
                    # Add column headers before new project
                    expanded_data.append(dict(zip(columns, columns)))
                    current_project = project_desc
                
                # Determine billing type based on Report Code
                report_code = str(row.get('Report Code', '')).lower()
                billing_type = ('Fixed' if any(term in report_code for term in fixed_terms) 
                              else 'T&E')
                if row.get('Activity Description'):
                    activities = [act.strip() for act in row['Activity Description'].split('|') 
                                if act.strip()]
                else:
                    activities = []
                
                # If no activities, add the row with project info
                if not activities:
                    estimated_rev = float(row.get('Estimated Revenue', 0))  # Change from Budget
                    actual_rev = float(row.get('Actual Revenue', 0))       # Change from Sales Amount
                    remaining = estimated_rev - actual_rev
                    
                    expanded_data.append({
                        'Project': row.get('Project', ''),
                        'Project Description': project_desc,
                        'Activity Sequence': row.get('Activity Sequence', ''),
                        'Activity': row.get('Activity', ''),  # Add this line
                        'Activity Description': '',
                        'Status': row.get('Status', ''),
                        'Billing Type': billing_type,
                        'Estimated Revenue': estimated_rev,
                        'Actual Revenue': actual_rev,
                        'Remaining Budget': remaining
                    })
                else:
                    # Add a row for each activity
                    for activity in sorted(set(activities)):
                        budgeted_revenue = float(row.get('Budget', 0))
                        actual_revenue = float(row.get('Sales Amount', 0))
                        budget_remaining = budgeted_revenue - actual_revenue
                        expanded_data.append({
                            'Project': row.get('Project', ''),
                            'Project Description': project_desc,
                            'Activity': row.get('Activity', ''),
                            'Activity Sequence': row.get('Activity Sequence', ''),  # Added missing Activity Sequence
                            'Activity Description': activity,
                            'Status': row.get('Status', ''),
                            'Invoice Quantity': row.get('Invoice Quantity', 0),
                            'Sales Quantity': row.get('Sales Quantity', 0),
                            'Report Code': row.get('Report Code', ''),
                            'Billing Type': billing_type,
                            'Budgeted Revenue': budgeted_revenue,
                            'Actual Revenue': actual_revenue,
                            'Budget Remaining': budget_remaining
                        })
            
            # Create DataFrame with specified columns
            expanded_df = pd.DataFrame(expanded_data, columns=columns)
            
            # Write to Excel with formatting
            expanded_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                na_rep=''
            )
            
            # Get worksheet reference
            worksheet = writer.sheets[sheet_name]
            
            # Set column widths and format headers
            column_widths = {
                'Project': 15,
                'Project Description': 50,
                'Activity Sequence': 15,
                'Activity': 30,
                'Activity Description': 100,
                'Status': 15,
                'Billing Type': 15,
                'Estimated Revenue': 15,
                'Actual Revenue': 15,
                'Remaining Budget': 15
            }
            
            # Format headers
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D3D3D3',
                'border': 1
            })
            
            # Apply column widths and header formatting
            for idx, col in enumerate(expanded_df.columns):
                worksheet.set_column(idx, idx, column_widths.get(col, 15))
                
                # Format all header rows
                for row_idx, row in expanded_df.iterrows():
                    if row['Project'] == 'Project':  # This is a header row
                        worksheet.write(row_idx + 1, idx, row[col], header_format)
            
            # Add freeze panes
            worksheet.freeze_panes(1, 0)
            
            logger.info(f"Successfully wrote sheet '{sheet_name}' with {len(expanded_df)} rows")
            
        except Exception as e:
            logger.error(f"Error writing sheet for {manager}: {str(e)}", exc_info=True)
            raise

    def _process_manager_data(self, manager: str, all_manager_data: List[pd.DataFrame]) -> pd.DataFrame:
        """Process and combine all data for a manager."""
        try:
            logger.info(f"\nProcessing data for {manager}")
            
            # Find P file data first (it has core project info)
            p_data = None
            other_data = []
            
            for df in all_manager_data:
                if ('Project' in df.columns and 'Manager Description' in df.columns and 
                    'Project Description' in df.columns):
                    p_data = df.copy()
                else:
                    other_data.append(df)
                    
            if p_data is None:
                logger.warning(f"No base project data found for {manager}")
                return pd.DataFrame()
                
            # Keep essential columns from PT data
            pt_data = None
            for df in other_data:
                if 'Sales Amount' in df.columns or 'Sales Quantity' in df.columns:
                    # Define possible columns we want to aggregate
                    possible_cols = {
                        'Sales Amount': 'sum',
                        'Sales Quantity': 'sum',
                        'Invoice Status': 'first',
                        'Status': 'first',
                        'Activity Description': lambda x: ' | '.join(x.dropna().unique()),
                        'Activity Sequence': lambda x: ' | '.join(str(s) for s in x.dropna().unique()),  # Changed to keep unique sequences
                        'Activity': 'first'
                    }
                    
                    # Only include columns that actually exist in the dataframe
                    agg_dict = {col: agg for col, agg in possible_cols.items() 
                              if col in df.columns}
                    
                    # Must have Project column for grouping
                    if 'Project' in df.columns:
                        pt_cols = ['Project'] + list(agg_dict.keys())
                        pt_data = df[pt_cols].copy()
                        logger.info(f"PT columns before grouping: {pt_cols}")
                        if 'Activity' in pt_data.columns:
                            logger.info(f"Activity values before grouping: {pt_data['Activity'].unique()}")
                        
                        # Group by Project without including it in agg_dict
                        pt_data = pt_data.groupby('Project', as_index=False).agg(agg_dict)
                        if 'Activity' in pt_data.columns:
                            logger.info(f"Activity values after grouping: {pt_data['Activity'].unique()}")
                        break
                        
            # Merge PT data if available
            if pt_data is not None:
                final_report = pd.merge(p_data, pt_data, on='Project', how='left')
                logger.info(f"Merged PT data - shape: {final_report.shape}")
            else:
                final_report = p_data.copy()
                logger.info(f"No PT data to merge - using P data only")
                       
            # Handle missing values
            numeric_cols = ['Sales Amount', 'Sales Quantity']
            for col in final_report.columns:
                if col in numeric_cols:
                    final_report[col] = pd.to_numeric(final_report[col], errors='coerce').fillna(0)
                else:
                    final_report[col] = final_report[col].fillna('')
                    
            # Add calculations if possible
            if all(col in final_report.columns for col in ['Sales Amount', 'Sales Quantity']):
                try:
                    mask = (final_report['Sales Quantity'] != 0)
                    final_report['Revenue per Unit'] = pd.NA
                    final_report.loc[mask, 'Revenue per Unit'] = (
                        final_report.loc[mask, 'Sales Amount'] /  
                        final_report.loc[mask, 'Sales Quantity']
                    )
                    logger.info("Calculated Revenue per Unit")
                except Exception as e:
                    logger.error(f"Error calculating Revenue per Unit: {str(e)}")
                    
            logger.info(f"Final shape: {final_report.shape}")
            logger.info(f"Final columns: {final_report.columns.tolist()}")
            
            return final_report

        except Exception as e:
            logger.error(f"Error processing manager data: {str(e)}", exc_info=True)
            return pd.DataFrame()

    def _format_project_activities(self, df: pd.DataFrame) -> pd.DataFrame:
        """Format dataframe with projects and activities in hierarchical order."""
        try:
            logger.info("Formatting project activities hierarchy...")
            
            # Sort by Project Description
            df = df.sort_values('Project Description')
            
            # Create new dataframe for formatted output
            formatted_rows = []
            current_project = None
            
            for _, row in df.iterrows():
                project_desc = row['Project Description']
                activities = row.get('Activity Description', '').split(' | ')
                
                # If this is a new project, add project info
                if project_desc != current_project:
                    # Add spacing between projects (3 empty rows)
                    if current_project is not None:
                        for _ in range(3):
                            formatted_rows.append({col: '' for col in df.columns})
                    
                    # Add project row
                    formatted_rows.append(row.to_dict())
                    current_project = project_desc
                    
                    # Add activity rows if they exist
                    if activities and activities[0]:  # Check if there are activities
                        for activity in sorted(set(activities)):  # Use set to get unique activities
                            activity_row = {col: '' for col in df.columns}
                            activity_row['Activity Description'] = f"  • {activity}"  # Indent activities
                            formatted_rows.append(activity_row)
                
                # If same project but new activities, append them
                else:
                    new_activities = set(activities) - set(
                        row['Activity Description'].replace('  • ', '')
                        for row in formatted_rows 
                        if row.get('Activity Description', '').startswith('  • ')
                    )
                    for activity in sorted(new_activities):
                        activity_row = {col: '' for col in df.columns}
                        activity_row['Activity Description'] = f"  • {activity}"
                        formatted_rows.append(activity_row)
            
            # Convert back to DataFrame
            formatted_df = pd.DataFrame(formatted_rows)
            
            logger.info(f"Formatted {len(df)} projects with activities")
            return formatted_df
            
        except Exception as e:
            logger.error(f"Error formatting project activities: {str(e)}")
            return df  # Return original dataframe if formatting fails

def main():
    """Main program execution."""
    print("Starting program...")
    
    # File paths configuration
    file_paths = {
        "A": "C:/Reporting/Py/A.xlsx",
        "AE": "C:/Reporting/Py/AE.xlsx",
        "PI": "C:/Reporting/Py/PI.xlsx",
        "PT": "C:/Reporting/Py/PT.xlsx",
        "SI": "C:/Reporting/Py/SI.xlsx",
        "P": "C:/Reporting/Py/P.xlsx"
    }
    
    print("Initializing ProjectAnalyzer...")
    analyzer = ProjectAnalyzer(file_paths)
    
    print("Attempting to load data...")
    if not analyzer.load_data():
        print("Failed to load data")
        return
    
    while True:
        print("\nProject Analysis Menu:")
        print("1. Generate Weekly Report")
        print("2. Refresh Data from Files")
        print("3. Inspect Data")
        print("4. Export Data for Viewing")
        print("5. Exit")
        print("6. Debug DataFrame Info")  # New option
        print("7. Clear Cache")  # New option
        
        try:
            choice = input("Enter your choice (1-7): ")
            print(f"You selected option: {choice}")
            
            if choice == "1":
                output_path = "C:/Reporting/Py/Weekly_Report.xlsx"
                print(f"Attempting to generate report at: {output_path}")
                if not analyzer.generate_weekly_report(output_path):
                    print("Failed to generate weekly report")
                else:
                    print("Report generated successfully")
            elif choice == "6":  # New option handler
                df_name = input("Enter dataframe name to debug (A, AE, PI, PT, SI, P): ")
                analyzer.debug_df_info(df_name)
            elif choice == "2":
                if analyzer.load_data(force_refresh=True):
                    print("Data refreshed successfully")
                else:
                    print("Failed to refresh data")
            elif choice == "3":
                analyzer.inspect_data()
            elif choice == "4":
                print("Export functionality not yet implemented")
                # analyzer.export_for_viewing()  # Comment out until implemented
            elif choice == "5":
                print("Exiting program...")
                break
            elif choice == "7":
                try:
                    os.remove("C:/Reporting/Py/data_cache.pkl")
                    print("Cache cleared successfully")
                except FileNotFoundError:
                    print("Cache already cleared")
            else:
                print("Invalid choice. Please enter 1-7.")
        except Exception as e:
            print(f"Error processing choice: {str(e)}")
            logger.error(f"Error in main loop: {str(e)}", exc_info=True)

if __name__ == "__main__":
    main()