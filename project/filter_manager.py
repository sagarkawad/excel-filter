import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import os

# Global variables
filter_rows = []
filters_frame = None


def add_filter_row():

    # ... (code to add a filter row)
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
    # ... (code to delete a filter row)
    filter_rows[:] = [(frame, col, val) for frame, col, val in filter_rows if frame != row_frame]
    # Destroy the row widgets
    row_frame.destroy()

def filter_and_save(df, original_file_path):
    # ... (code to filter and save the DataFrame)
    filtered_df = df.copy()
    filename_parts = []
    
    print("Initial DataFrame shape:", filtered_df.shape)
    
    for _, column_entry, value_entry in filter_rows:
        column_name = column_entry.get().strip().lower()
        filter_value = value_entry.get().strip().lower()

        print(f"\nProcessing filter: {column_name} contains {filter_value}")
        
        if not column_name or not filter_value:
            continue
            
        if column_name not in df.columns.str.lower():
            messagebox.showerror("Error", f"Column '{column_name}' not found in the Excel file!")
            return False
        
        # Determine column type and convert filter value accordingly
        column_type = str(filtered_df[column_name].dtype)
        print(f"Column '{column_name}' type:", column_type)
        
        try:
            if 'int' in column_type or 'float' in column_type:
                # Handle numeric columns - check for exact match
                filter_value = float(filter_value)
                filtered_df = filtered_df[filtered_df[column_name] == filter_value]  # Exact match
                
            else:
                # Handle string columns - use str.contains for substring matching
                filtered_df = filtered_df[filtered_df[column_name].astype(str).str.contains(str(filter_value), case=False, na=False)]
                
            
            print(f"Rows remaining after this filter: {len(filtered_df)}")
            
            if filtered_df.empty:
                messagebox.showerror("Error", f"No matches found after applying filter: {column_name} contains {filter_value}")
                return False
                
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid value format for column {column_name}: {str(e)}")
            return False
        
        filename_parts.append(f"{column_name}_{filter_value}")
    
    if filtered_df.empty:
        messagebox.showerror("Error", "No rows found matching all filter criteria")
        return False
    
    # Ask user for the location to save the file
    save_file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")],
        initialfile="_".join(filename_parts) + "_filtered.xlsx"
    )
    
    if save_file_path:
        # Save filtered dataframe and show preview
        filtered_df.to_excel(save_file_path, index=False)
    
        # Show preview of filtered data
        preview_message = f"Filtered results ({len(filtered_df)} rows):\n\n"
        preview_message += filtered_df.head().to_string()
        messagebox.showinfo("Success", f"Filtered file saved as: {os.path.basename(save_file_path)}\n\n{preview_message}")
        return True
    else:
        messagebox.showerror("Error", "File save operation was cancelled.")
        return False

def check_filters():
    # ... (code to check filters)
    global tree, file_path, row_count_label  # Declare row_count_label as global
    if not file_path:
        return  # No file loaded, exit the function

    try:
        df = pd.read_excel(file_path)  # Reload the DataFrame from the file
        filtered_df = df.copy()

        for _, column_entry, value_entry in filter_rows:
            column_name = column_entry.get().strip().lower()
            filter_value = value_entry.get().strip().lower()

            if not column_name or not filter_value:
                continue
            
            lower_columns = [col.lower() for col in df.columns]  # Convert df.columns to lowercase
            if column_name not in lower_columns:
                messagebox.showerror("Error", f"Column '{column_name}' not found in the Excel file!")
                return
            
            actual_column_name = df.columns[lower_columns.index(column_name)]  # Get the actual column name

            # Determine column type
            column_type = str(filtered_df[actual_column_name].dtype)
            
            # Filter logic
            if 'int' in column_type or 'float' in column_type:
                # Handle numeric columns - check for exact match
                try:
                    filtered_df = filtered_df[filtered_df[actual_column_name] == float(filter_value)]  # Exact match
                except ValueError:
                    continue  # Skip if conversion fails
            else:
                # Handle string columns - use str.contains for substring matching
                filtered_df = filtered_df[filtered_df[actual_column_name].astype(str).str.contains(filter_value, case=False, na=False)]

        # Clear the Treeview
        for item in tree.get_children():
            tree.delete(item)

        # Insert the filtered data into the Treeview
        for index, row in filtered_df.iterrows():
            tree.insert("", "end", values=list(row))

        # Display the number of filtered rows
        row_count_label.config(text=f"Total number of rows: {len(df)} | Filtered rows: {len(filtered_df)}")

    except Exception as e:
        messagebox.showerror("Error", f"Error filtering data: {str(e)}")

