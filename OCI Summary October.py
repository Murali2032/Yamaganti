import os
import pandas as pd

# Define the paths
excel_file_path = r"C:\Users\YamagantiMuraliKrish\Downloads\File Names to replace.xlsx"  # Updated Excel file path
file_folder_path = r"C:\Users\YamagantiMuraliKrish\Music\October 2024"  # Path to the folder containing the files

# Load the Excel file specifying the sheet name
try:
    # Load the sheet named "Summary" (Sheet 2)
    df = pd.read_excel(excel_file_path, sheet_name='Summary', engine='openpyxl')

    # Check if required columns exist
    if 'Orginal File name' not in df.columns or 'Rename File to' not in df.columns:
        raise ValueError("Excel file must contain columns 'Orginal File name' and 'Rename File to'.")

    updated_files = []  # To store the updated file names

    # Loop through the DataFrame rows
    for index, row in df.iterrows():
        old_file_name = str(row['Orginal File name'])  # Read file name from column Orginal File name
        new_file_name = str(row['Rename File to'])  # Read new file name from column Rename File to

        old_file_path = os.path.join(file_folder_path, old_file_name)
        new_file_path = os.path.join(file_folder_path, new_file_name)

        # Check if the old file exists
        if os.path.exists(old_file_path):
            # Rename the file
            os.rename(old_file_path, new_file_path)
            updated_files.append(new_file_name)  # Store the new file name
            print(f"Renamed: {old_file_name} -> {new_file_name}")
        else:
            print(f"File not found: {old_file_name}")

    # Print all updated file names
    print("\nUpdated File Names:")
    for file in updated_files:
        print(file)

except Exception as e:
    print(f"An error occurred: {e}")