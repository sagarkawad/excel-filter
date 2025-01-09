import os
import pandas as pd
from tkinter import filedialog, messagebox
import subprocess
import sys

def open_excel_file(file_path):
    if os.name == 'nt':  # For Windows
        os.startfile(file_path)
    elif os.name == 'posix':  # For macOS and Linux
        if sys.platform == 'darwin':  # macOS
            subprocess.call(['open', file_path])
        else:  # Linux
            subprocess.call(['xdg-open', file_path])

def load_excel():
    global tree, file_path, row_count_label  # Declare row_count_label as global
    file_path = filedialog.askopenfilename(
        title="Open Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if file_path:
        try:
            df = pd.read_excel(file_path)
            # ... (rest of the load_excel function code)
            return df, file_path
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
