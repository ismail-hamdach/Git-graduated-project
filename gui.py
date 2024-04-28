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

    # Load the Excel files into Pandas DataFrames
    dfs = [pd.read_excel(os.path.join(current_directory, file)) for file in file_paths]

    # Define the value you want to search for in the "contrat" column
    search_value = entry_search_value.get()

    # Convert search value to integer if possible
    try:
        search_value = int(search_value)
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter a valid integer search value.")
        return

    # Iterate through each DataFrame to search for the value
    found_file = None
    for i, df in enumerate(dfs):
        if search_value in df["Contrat"].values:
            found_file = file_paths[i]
            break

    if found_file:
        messagebox.showinfo("File Found", f"Value found in {found_file}")
        os.startfile(os.path.join(current_directory, found_file))  # Open the file using the default program
    else:
        messagebox.showinfo("Value Not Found", "Value not found in any file.")

print("Hello world")

# Create the main window
root = tk.Tk()
root.title("Excel File Search and Open")

# Create a label and entry for the search value
label_search_value = tk.Label(root, text="Search Value:")
label_search_value.grid(row=0, column=0, padx=5, pady=5)
entry_search_value = tk.Entry(root)
entry_search_value.grid(row=0, column=1, padx=5, pady=5)

# Create a button to search and open Excel files
button_search = tk.Button(root, text="Search and Open Excel Files", command=search_and_open_excel)
button_search.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

# Run the Tkinter event loop
root.mainloop()
