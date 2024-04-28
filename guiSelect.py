import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def search_and_open_excel():
    # Get the current directory
    current_directory = os.getcwd()

    # Get a list of all files in the current directory
    file_paths = [file for file in os.listdir(current_directory) if file.endswith(".xlsx")]

    if not file_paths:
        messagebox.showinfo("No Excel Files", "No Excel files found in the current directory.")
        return

    # Load the first Excel file into a Pandas DataFrame to get column names
    first_file_path = os.path.join(current_directory, file_paths[0])
    first_df = pd.read_excel(first_file_path)

    # Get column names from the first DataFrame
    column_names = first_df.columns.tolist()

    # Create a dropdown menu for column selection
    selected_column = column_selection.get()

    # Check if a column is selected
    if not selected_column:
        messagebox.showinfo("No Column Selected", "Please select a column to search.")
        return

    # Load all Excel files into Pandas DataFrames
    dfs = [pd.read_excel(os.path.join(current_directory, file)) for file in file_paths]

    # Define the value you want to search for
    search_value = entry_search_value.get()

    # Iterate through each DataFrame to search for the value in the selected column
    found_file = None
    for i, df in enumerate(dfs):
        if search_value in df[selected_column].values:
            found_file = file_paths[i]
            break

    if found_file:
        messagebox.showinfo("File Found", f"Value found in {found_file}")
        os.startfile(os.path.join(current_directory, found_file))  # Open the file using the default program
    else:
        messagebox.showinfo("Value Not Found", "Value not found in any file.")

# Create the main window
root = tk.Tk()
root.title("Excel File Search and Open")

# Create a label and entry for the search value
label_search_value = tk.Label(root, text="Search Value:")
label_search_value.grid(row=0, column=0, padx=5, pady=5)
entry_search_value = tk.Entry(root)
entry_search_value.grid(row=0, column=1, padx=5, pady=5)

# Create a label and dropdown menu for column selection
label_column_select = tk.Label(root, text="Select Column:")
label_column_select.grid(row=1, column=0, padx=5, pady=5)
column_selection = tk.StringVar(root)
column_selection.set("")  # Set default value to blank
dropdown_column_select = tk.OptionMenu(root, column_selection, "")
dropdown_column_select.grid(row=1, column=1, padx=5, pady=5)

# Function to update the column dropdown menu based on the first file's columns
def update_column_dropdown():
    # Get the first Excel file in the current directory
    current_directory = os.getcwd()
    file_paths = [file for file in os.listdir(current_directory) if file.endswith(".xlsx")]

    if file_paths:
        first_file_path = os.path.join(current_directory, file_paths[0])
        first_df = pd.read_excel(first_file_path)
        column_names = first_df.columns.tolist()
        dropdown_column_select["menu"].delete(0, "end")  # Clear the menu
        for column in column_names:
            dropdown_column_select["menu"].add_command(label=column, command=lambda value=column: column_selection.set(value))

# Update the column dropdown menu
update_column_dropdown()

# Create a button to search and open Excel files
button_search = tk.Button(root, text="Search and Open Excel Files", command=search_and_open_excel)
button_search.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Run the Tkinter event loop
root.mainloop()
