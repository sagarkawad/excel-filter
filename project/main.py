import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from PIL import Image, ImageTk  # Import Pillow for image handling
import sys
import subprocess

# Create the main application window
root = tk.Tk()
root.title("Load Excel File Example")
root.attributes('-fullscreen', True)  # Set the window to full screen
root.geometry("400x300")

# Create a canvas to hold the background image
canvas = tk.Canvas(root, width=100, height=100)
canvas.pack(fill="none", expand=False)

# Load the logo image
logo_path = "assets/icon/navbarLogo.jpeg"
if os.path.exists(logo_path):
    logo_image = Image.open(logo_path)  # Use Pillow to open the image
    logo_image = logo_image.resize((100, 100), Image.LANCZOS)  # Use LANCZOS instead of ANTIALIAS
    logo_image = ImageTk.PhotoImage(logo_image)  # Convert to PhotoImage for Tkinter
    # Create an image on the canvas
    canvas.create_image(0, 0, anchor=tk.NW, image=logo_image)
else:
    messagebox.showerror("Error", f"Logo image not found at: {logo_path}")
    logo_image = None  # Set to None or a default image if needed

# Create a frame to hold all filter inputs
filters_frame = tk.Frame(root, bg='white')  # Set background color for the frame
filters_frame.pack(pady=10)  # Use pack instead of place to ensure visibility

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

# Function to open the Excel file
def open_excel_file(file_path):
    if os.name == 'nt':  # For Windows
        os.startfile(file_path)
    elif os.name == 'posix':  # For macOS and Linux
        if sys.platform == 'darwin':  # macOS
            subprocess.call(['open', file_path])
        else:  # Linux
            subprocess.call(['xdg-open', file_path])

# Function to handle double-click event on the Treeview
def on_treeview_double_click(event):
    item = tree.selection()  # Get the selected item
    if item:  # Check if an item is selected
        open_excel_file(file_path)  # Open the Excel file

def load_excel():
    global tree, file_path, row_count_label  # Declare row_count_label as global
    file_path = filedialog.askopenfilename(
        title="Open Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if file_path:
        try:
            df = pd.read_excel(file_path)
            
            # Clear existing UI elements except the filters_frame and load_button
            for widget in root.winfo_children():
                if widget != filters_frame and widget != load_button and isinstance(widget, (tk.Label, tk.Text, tk.Button, tk.Frame, ttk.Treeview)):
                    widget.destroy()
            
            # Clear existing filter rows
            for frame, _, _ in filter_rows:
                frame.destroy()
            filter_rows.clear()
            
            success_label = tk.Label(root, text=f"Excel file loaded successfully: {os.path.basename(file_path)}")
            success_label.pack(pady=10)
            
            # Create a frame to hold the Treeview and scrollbars
            tree_frame = tk.Frame(root)
            tree_frame.pack(expand=True, fill='both', pady=10)
            
            # Create a style for the Treeview
            style = ttk.Style()
            style.configure("Treeview", font=('TkDefaultFont', 10))  # Set font to default
            style.configure("Treeview.Heading", font=('TkDefaultFont', 10, 'bold'))  # Set font for headings

            # Add border to the Treeview cells
            style.configure("Treeview", bordercolor="black", borderwidth=1)  # Set border color and width

            # Create a Treeview widget to display the DataFrame
            tree = ttk.Treeview(tree_frame, columns=list(df.columns), show='headings')
            tree.pack(side=tk.LEFT, expand=True, fill='both')

            # Set the column headings and widths
            for col in df.columns:
                tree.heading(col, text=col, anchor='center')  # Center the column headings
                tree.column(col, anchor='center', width=100)  # Set a default width for columns

            # Add horizontal and vertical scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            vsb.pack(side=tk.RIGHT, fill='y')
            tree.configure(yscrollcommand=vsb.set)

            hsb = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
            hsb.pack(fill='x')
            tree.configure(xscrollcommand=hsb.set)

            # Set the column headings
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor='center', width=100)  # Set a default width for columns
            tree['show'] = 'headings'  # Show only headings, no empty first column
            tree.configure(selectmode='browse')  # Allow single selection
            
            # Add grid lines
            tree.tag_configure('grid')
            for i in range(len(df.columns)):
                tree.column(i, width=100)  # Set width for each column
            
            # Insert the data into the Treeview
            for index, row in df.iterrows():
                tree.insert("", "end", values=list(row))
            
            # Display the total number of rows
            row_count_label = tk.Label(root, text=f"Total number of rows: {len(df)}")
            row_count_label.pack(pady=5)
            
            # Add "Add Filter" button before adding the first row
            add_filter_btn = tk.Button(root, text="Add Filter", command=add_filter_row)
            add_filter_btn.pack(pady=5)
            
            # Add first filter row by default
            add_filter_row()

            # Add "Check Filter" button at the bottom of the Excel file display
            check_filter_btn = tk.Button(root, text="Check Filter", command=check_filters)
            check_filter_btn.pack(pady=5)
            
            filter_button = tk.Button(root, text="Filter and Save", 
                                    command=lambda: filter_and_save(df, file_path))
            filter_button.pack(pady=5)

            # Bind double-click event to the Treeview
            tree.bind("<Double-1>", on_treeview_double_click)  # Bind double-click event
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

def filter_displayed_rows():
    global tree, file_path  # Ensure tree and file_path are accessible
    if not file_path:
        return  # No file loaded, exit the function

    try:
        df = pd.read_excel(file_path)  # Reload the DataFrame from the file
        filtered_df = df.copy()

        for _, column_entry, value_entry in filter_rows:
            column_name = column_entry.get().strip().lower()  # Convert to lowercase
            filter_value = value_entry.get().strip()

            if not column_name or not filter_value:
                continue
            
            # Ensure case-insensitive check for column names
            lower_columns = [col.lower() for col in df.columns]  # Convert df.columns to lowercase
            if column_name not in lower_columns:
                messagebox.showerror("Error", f"Column '{column_name}' not found in the Excel file!")
                return
            
            # Get the actual column name from the original DataFrame
            actual_column_name = df.columns[lower_columns.index(column_name)]

            # Filter logic
            filtered_df = filtered_df[filtered_df[actual_column_name].astype(str).str.contains(filter_value, case=False, na=False)]

        # Clear the Treeview
        for item in tree.get_children():
            tree.delete(item)

        # Insert the filtered data into the Treeview
        for index, row in filtered_df.iterrows():
            tree.insert("", "end", values=list(row))

    except Exception as e:
        messagebox.showerror("Error", f"Error filtering data: {str(e)}")

# New function to check and display filtered results
def check_filters():
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

# Function to handle zooming in and out
def zoom(event):
    global tree  # Ensure tree is accessible
    current_font = tree.cget("font")
    font_size = int(current_font.split()[1])  # Get current font size

    if event.delta > 0:  # Zoom in
        new_font_size = font_size + 2
    else:  # Zoom out
        new_font_size = max(8, font_size - 2)  # Prevent font size from going below 8

    # Update the font size for the Treeview and its headings
    new_font = (current_font.split()[0], new_font_size)
    tree.configure(font=new_font)
    style.configure("Treeview", font=new_font)  # Update Treeview style

# Bind the zoom function to the mouse wheel event
root.bind("<Control-MouseWheel>", zoom)

# Button to load an Excel file
load_button = tk.Button(root, text="Load Excel File", command=load_excel)
load_button.pack(pady=10)

# Start the main event loop
root.mainloop()
