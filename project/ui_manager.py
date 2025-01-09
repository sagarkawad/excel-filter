import tkinter as tk
from tkinter import messagebox, ttk
import os
from PIL import Image, ImageTk 
from file_loader import load_excel, open_excel_file



def setup_ui(root):
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

         # Function to handle double-click event on the Treeview
def on_treeview_double_click(event, tree, file_path):
    item = tree.selection()  # Get the selected item
    if item:  # Check if an item is selected
        open_excel_file(file_path)  # Open the Excel file

def create_treeview(root, filters_frame, filter_rows, filter_and_save, add_filter_row, check_filters):
        try:
            df, file_path = load_excel()
            
            # Button to load an Excel file
            load_button = tk.Button(root, text="Load Excel File", command=load_excel)
            load_button.pack(pady=10)
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
            
            # Create a style for the Treeview
            style = ttk.Style()
            style.configure("Treeview", font=('TkDefaultFont', 14))  # Increased font size to 14
            style.configure("Treeview.Heading", font=('TkDefaultFont', 14, 'bold'))  # Increased font size for headings

            # Add border to the Treeview cells
            style.configure("Treeview", bordercolor="black", borderwidth=1)  # Set border color and width
            style.map("Treeview", bordercolor=[('selected', 'blue')])  # Optional: Change border color when selected

            # Create a frame to hold the Treeview and scrollbars
            tree_frame = tk.Frame(root)
            tree_frame.pack(expand=True, fill='both', pady=10)
            
            # Create a Treeview widget to display the DataFrame
            tree = ttk.Treeview(tree_frame, columns=list(df.columns), show='headings')
            tree.pack(side=tk.LEFT, expand=True, fill='both')

           # Ensure df is defined and has columns
            if df is not None and not df.empty:
                for col in df.columns:
                    tree.heading(col, text=col, anchor='center')  # Center the column headings
                    tree.column(col, anchor='center', width=100)  # Set width to 300 pixels

            # Add horizontal and vertical scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            vsb.pack(side=tk.RIGHT, fill='y')
            tree.configure(yscrollcommand=vsb.set)

            hsb = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
            hsb.pack(fill='x')
            tree.configure(xscrollcommand=hsb.set)

            # Add grid lines and borders for each cell
            tree.tag_configure('grid')  # Set border width and style
            for i in range(len(df.columns)):
                
                tree.column(i, width=300)  # Ensure width is set to 300 pixels for each column

            # Insert the data into the Treeview
            for index, row in df.iterrows():
                tree.insert("", "end", values=list(row))  # Apply grid tag for borders
            
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
            tree.bind("<Double-1>", lambda event: on_treeview_double_click(event, tree, file_path))  # Bind double-click event with lambda
            
            # Determine the maximum width for each column based on the content
            for i in range(len(df.columns)):
                max_width = max(df[df.columns[i]].astype(str).map(len).max(), len(df.columns[i]))  # Get max width of content
                tree.column(i, width=(max_width * 10) + 10)  # Set width based on content length (multiplied for better visibility)

        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")