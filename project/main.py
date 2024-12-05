import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Create the main application window
root = tk.Tk()
root.title("Load Excel File Example")
root.geometry("400x300")

# Create a frame to hold all filter inputs
filters_frame = tk.Frame(root)
filters_frame.pack(pady=10)

# List to keep track of filter input rows
filter_rows = []

# Function to create a new filter input row
def add_filter_row():
    row_frame = tk.Frame(filters_frame)
    row_frame.pack(pady=2)
    
    # Label and Entry for column
    label1 = tk.Label(row_frame, text="Column:")
    label1.pack(side=tk.LEFT, padx=5)
    input_box1 = tk.Entry(row_frame)
    input_box1.pack(side=tk.LEFT, padx=5)
    
    # Label and Entry for value
    label2 = tk.Label(row_frame, text="Value:")
    label2.pack(side=tk.LEFT, padx=5)
    input_box2 = tk.Entry(row_frame)
    input_box2.pack(side=tk.LEFT, padx=5)
    
    # Delete button for this row
    delete_btn = tk.Button(row_frame, text="X", command=lambda: delete_filter_row(row_frame))
    delete_btn.pack(side=tk.LEFT, padx=5)
    
    # Add row components to list
    filter_rows.append((row_frame, input_box1, input_box2))

def delete_filter_row(row_frame):
    # Remove row from filter_rows list
    filter_rows[:] = [(frame, col, val) for frame, col, val in filter_rows if frame != row_frame]
    # Destroy the row widgets
    row_frame.destroy()

# Modified filter_and_save function to handle multiple filters
def filter_and_save(df, original_file_path):
    filtered_df = df.copy()
    filename_parts = []
    
    print("Initial DataFrame shape:", filtered_df.shape)
    
    for _, column_entry, value_entry in filter_rows:
        column_name = column_entry.get().strip()
        filter_value = value_entry.get().strip()

        print(f"\nProcessing filter: {column_name} = {filter_value}")
        
        if not column_name or not filter_value:
            continue
            
        if column_name not in df.columns:
            messagebox.showerror("Error", f"Column '{column_name}' not found in the Excel file!")
            return False
        
        # Determine column type and convert filter value accordingly
        column_type = str(filtered_df[column_name].dtype)
        print(f"Column '{column_name}' type:", column_type)
        
        try:
            if 'int' in column_type or 'float' in column_type:
                # Handle numeric columns
                filter_value = float(filter_value)
                filtered_df = filtered_df[filtered_df[column_name] == filter_value]
            else:
                # Handle string columns - convert both column and filter value to strings
                filtered_df = filtered_df[filtered_df[column_name].astype(str).str.lower() == str(filter_value).lower()]
            
            print(f"Rows remaining after this filter: {len(filtered_df)}")
            
            if filtered_df.empty:
                messagebox.showerror("Error", f"No matches found after applying filter: {column_name} = {filter_value}")
                return False
                
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid value format for column {column_name}: {str(e)}")
            return False
        
        filename_parts.append(f"{column_name}_{filter_value}")
    
    if filtered_df.empty:
        messagebox.showerror("Error", "No rows found matching all filter criteria")
        return False
    
    # Create new filename
    file_dir = os.path.dirname(original_file_path)
    if len(filename_parts) != 0:
        file_name = "_".join(filename_parts)
        new_file_path = os.path.join(file_dir, f"{file_name}_filtered.xlsx")
        # Save filtered dataframe and show preview
        filtered_df.to_excel(new_file_path, index=False)
    
        # Show preview of filtered data
        preview_message = f"Filtered results ({len(filtered_df)} rows):\n\n"
        preview_message += filtered_df.head().to_string()
        messagebox.showinfo("Success", f"Filtered file saved as: {os.path.basename(new_file_path)}\n\n{preview_message}")
        return True
    else:
        messagebox.showerror("Error", "Please add atleast one filter to continue!")


# Modified load_excel function
def load_excel():
    file_path = filedialog.askopenfilename(
        title="Open Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if file_path:
        try:
            df = pd.read_excel(file_path)
            
            # Clear existing UI elements except the filters_frame and load_button
            for widget in root.winfo_children():
                if widget != filters_frame and widget != load_button and isinstance(widget, (tk.Label, tk.Text, tk.Button, tk.Frame)):
                    widget.destroy()
            
            # Clear existing filter rows
            for frame, _, _ in filter_rows:
                frame.destroy()
            filter_rows.clear()
            
            success_label = tk.Label(root, text=f"Excel file loaded successfully: {os.path.basename(file_path)}")
            success_label.pack(pady=10)
            
            text_widget = tk.Text(root, height=20, width=80)
            text_widget.pack(pady=10)
            
            text_widget.insert(tk.END, df.to_string())
            text_widget.configure(state='disabled')  # Make the text widget read-only
            
            # Add "Add Filter" button before adding the first row
            add_filter_btn = tk.Button(root, text="Add Filter", command=add_filter_row)
            add_filter_btn.pack(pady=5)
            
            # Add first filter row by default
            add_filter_row()
            
            filter_button = tk.Button(root, text="Filter and Save", 
                                    command=lambda: filter_and_save(df, file_path))
            filter_button.pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

# Button to load an Excel file
load_button = tk.Button(root, text="Load Excel File", command=load_excel)
load_button.pack(pady=10)

# Start the main event loop
root.mainloop()
