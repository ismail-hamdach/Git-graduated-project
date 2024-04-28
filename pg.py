import os
import pandas as pd

# Get the current directory
current_directory = os.getcwd()

# Get a list of all files in the current directory
file_paths = [file for file in os.listdir(current_directory) if file.endswith(".xlsx")]
file_paths

# Load the Excel files into Pandas DataFrames
dfs = [pd.read_excel(os.path.join(current_directory, file)) for file in file_paths]


# Define the value you want to search for in the "contrat" column
search_value = 4598724

# Iterate through each DataFrame to search for the value
for i, df in enumerate(dfs):
    if search_value in df["Contrat"].values:
        print(f"Value found in {file_paths[i]}")
        # Open the file or perform any further actions here
        os.startfile(os.path.join(current_directory, file_paths[i]))  # Opens the file using the default program
        break
else:
    print("Value not found in any file.")
