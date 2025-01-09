import tkinter as tk
from ui_manager import setup_ui, create_treeview
from filter_manager import filter_and_save, add_filter_row, check_filters, filter_rows, filters_frame



# Function to create the main application window and UI components
def create_main_window():
    global filters_frame

    # Create the main application window
    root = tk.Tk()
    root.title("Load Excel File Example")
    root.attributes('-fullscreen', True)  # Set the window to full screen
    root.geometry("400x300")

     # Create a frame to hold all filter inputs
    filters_frame = tk.Frame(root, bg='white')  # Set background color for the frame
    filters_frame.pack(pady=10)  # Use pack instead of place to ensure visibility


    # Setup UI components
    setup_ui(root)
    create_treeview(root, filters_frame, filter_rows, filter_and_save, add_filter_row, check_filters)

    # # Button to load an Excel file
    # load_button = tk.Button(root, text="Load Excel File", command=load_excel)
    # load_button.pack(pady=10)

    # Start the main event loop
    root.mainloop()

# Call the function to create the main window
if __name__ == "__main__":
    create_main_window()
