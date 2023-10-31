import os
import pandas as pd
import re
import glob

# Get the directory where the script is located
script_directory = os.path.dirname(os.path.abspath(__file__))

# Set the folder path to be the same as the script directory
folder_path = script_directory

# Define a function to process a single Excel file


def process_excel_file(file_path):
    # Read the Excel file
    data_xls = pd.read_excel(file_path, index_col=None)

    def extract_custom(value):
        match = re.search(r'\b1\d{11}\b', str(value))
        # match = re.search(r'1\d{6}(?!(00010))\d{5}', str(value))
        # match = re.search(r'1\d{11}', str(value))
        return match.group() if match else None

    # Apply the custom function to the "Unnamed: 20" column
    data_xls["Unnamed: 20"] = data_xls["Unnamed: 20"].apply(extract_custom)

    # Convert the "Unnamed: 1" column to datetime format
    data_xls["Unnamed: 1"] = pd.to_datetime(
        data_xls["Unnamed: 1"], errors='coerce').dt.date

    # Convert the "Unnamed: 13" column to numeric (including handling non-numeric values)
    data_xls["Unnamed: 13"] = pd.to_numeric(
        data_xls["Unnamed: 13"], errors='coerce')

    # Format the numbers in the "Unnamed: 13" column to have two decimal places
    data_xls["Unnamed: 13"] = data_xls["Unnamed: 13"].round(2)

    # Filter rows where "Unnamed: 20" contains 11 digits starting with "1" and not ending with "00001"
    data_xls = data_xls[data_xls["Unnamed: 20"].notna()]

    data_xls = data_xls[data_xls["Unnamed: 13"].notna()]

    data_xls = data_xls[data_xls["Unnamed: 1"].notna()]

    selected_columns = ["Unnamed: 1", "Unnamed: 13", "Unnamed: 20"]
    result_df = data_xls[selected_columns]

    # Save the result as a CSV with semicolon as the separator
    result_df.to_csv('your_csv.csv', encoding='utf-8',
                     index=False, sep=';', header=False)


# Use glob to find all Excel files (with .xlsx extension) in the folder
xlsx_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Iterate through the Excel files and process each one
for file_path in xlsx_files:
    process_excel_file(file_path)
