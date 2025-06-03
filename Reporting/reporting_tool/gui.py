import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import os
import threading
from calculations import perform_calculations
from excel_output import write_to_excel, folder_path

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Report Generator")
        self.root.geometry("600x400")

        # Input folder selection
        self.input_folder = folder_path  # Default folder
        tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.folder_var = tk.StringVar(value=self.input_folder)
        tk.Entry(root, textvariable=self.folder_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)

        # Log display
        tk.Label(root, text="Logs:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.log_text = scrolledtext.ScrolledText(root, width=70, height=15)
        self.log_text.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

        # Buttons
        tk.Button(root, text="Run Report", command=self.run_report).grid(row=3, column=0, padx=5, pady=10)
        self.open_btn = tk.Button(root, text="Open Report", command=self.open_report, state=tk.DISABLED)
        self.open_btn.grid(row=3, column=1, padx=5, pady=10)

        # Redirect print to log_text
        self.original_print = print
        self.print = self.log_print

    def log_print(self, *args, **kwargs):
        """Redirect print statements to the GUI log area."""
        text = ' '.join(map(str, args))
        self.log_text.insert(tk.END, text + '\n')
        self.log_text.see(tk.END)
        self.root.update()
        self.original_print(*args, **kwargs)

    def browse_folder(self):
        """Open a file dialog to select the input folder."""
        folder = filedialog.askdirectory(initialdir=self.input_folder)
        if folder:
            self.input_folder = folder
            self.folder_var.set(folder)

    def run_report(self):
        """Run the report generation in a separate thread to keep GUI responsive."""
        self.log_text.delete(1.0, tk.END)  # Clear previous logs
        self.open_btn.config(state=tk.DISABLED)
        
        # Update the folder path in excel_output
        global folder_path
        folder_path = self.input_folder
        
        # Run the pipeline in a separate thread
        threading.Thread(target=self.process_report, daemon=True).start()

    def process_report(self):
        """Execute the report generation pipeline."""
        try:
            # Perform calculations
            final_report, employee_hours = perform_calculations()
            
            # Write to Excel
            write_to_excel(final_report, employee_hours)
            
            # Enable the "Open Report" button
            self.open_btn.config(state=tk.NORMAL)
        except Exception as e:
            self.print(f"Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def open_report(self):
        """Open the generated Excel report."""
        output_file = os.path.join(self.input_folder, "reportX.xlsx")
        if os.path.exists(output_file):
            os.startfile(output_file)  # Windows-specific; use 'open' on macOS, 'xdg-open' on Linux
        else:
            messagebox.showerror("Error", "Report file not found!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
```

### How to Use

1. **Save the Modules**:
   - Ensure `data_pull.py`, `calculations.py`, and `excel_output.py` are in `c:/Reporting/Py/Program/` (as provided in the previous response).
   - Save `gui.py` to `c:/Reporting/Py/Program/gui.py`.

2. **Dependencies**:
   - Tkinter is included with Python, so no additional installation is needed.
   - Ensure `pandas` and `openpyxl` are installed:
     ```powershell
     C:/Python313/python.exe -m pip install pandas openpyxl
     ```

3. **Run the GUI**:
   - Run:
     ```powershell
     & C:/Python313/python.exe c:/Reporting/Py/Program/gui.py
     ```
   - A window will appear with:
     - An input folder field (default: `C:\Reporting\Data Downloaded from IFS`) and a "Browse" button to change it.
     - A log area showing progress (e.g., "Found 1 AE files", "Report created").
     - A "Run Report" button to execute the pipeline.
     - An "Open Report" button (enabled after the report is generated) to open `reportX.xlsx`.

4. **Interact**:
   - Optionally select a different input folder using the "Browse" button.
   - Click "Run Report" to generate the report.
   - Click "Open Report" to view the generated `reportX.xlsx`.

### GUI Features
- **Input Folder Selection**: Users can browse for a folder instead of hardcoding the path.
- **Log Display**: Redirects all `print` statements to a scrollable text area in the GUI.
- **Responsive Design**: Uses threading to keep the GUI responsive during long tasks (e.g., file reading, Excel writing).
- **Open Report**: Provides a button to open the generated Excel file directly.
- **Error Handling**: Shows errors in a pop-up dialog if something goes wrong.

### Output
- **GUI Window**: Displays logs and allows interaction.
- **`reportX.xlsx`**: Generated in the selected folder (default: `C:\Reporting\Data Downloaded from IFS`), same as before:
  - Sheets: `Activity Report`, manager-specific tabs, `Employee Hours`.
  - Formatting: Bold headers, currency formats, green data bars for `Budget Remaining`.

### Considerations
- **Platform**: The `os.startfile` in `open_report` is Windows-specific. For macOS, use `os.system("open " + output_file)`; for Linux, use `os.system("xdg-open " + output_file)`.
- **Scalability**: This GUI is minimal. If you need more features (e.g., DataFrame previews, progress bars, more settings), we can expand it using a framework like PyQt for better table displays.
- **Performance**: Threading ensures the GUI remains responsive, but for very large datasets, you might need additional optimizations (e.g., background workers with progress updates).

### Does This Make Things Easier?
- **For Users**: Yes, significantly. They can interact visually without terminal knowledge, select folders easily, and open the report with a click.
- **For Development**: It adds a small layer of complexity (GUI code, threading), but Tkinter is simple, and the modular structure keeps it manageable.
- **For Maintenance**: The GUI is isolated in `gui.py`, so it doesn’t affect the core logic (`data_pull`, `calculations`, `excel_output`).

### Date and Time Context
It’s **04:31 PM AEST on Monday, May 19, 2025**. If you’d like to incorporate this into the GUI (e.g., timestamp the report filename as `reportX_20250519.xlsx` or display the date in the GUI), I can modify `excel_output.py` or `gui.py`. For example:
- Modify `excel_output.py` to name the file with the date:
  ```python
  from datetime import datetime
  output_file = os.path.join(folder_path, f"reportX_{datetime.now().strftime('%Y%m%d')}.xlsx")
  ```

### Next Steps
- **Confirm Functionality**: Run `gui.py` and verify it meets your needs.
- **Enhancements**: If you want additional GUI features (e.g., DataFrame previews, progress bar, timestamp in filename), let me know.
- **Alternative Frameworks**: If you prefer a more modern GUI (e.g., with better table displays), I can switch to PyQt or PySide, though they require installation.
- **Packaging**: To make it a standalone app, I can guide you on using PyInstaller to create an `.exe` for Windows.

Does this GUI approach make things easier for your use case? Let me know if you’d like adjustments!