import os
import pandas as pd

# Define the folder path and the target Shipment ID
folder_path = r'E:\shipment'  # Updated to your folder path
target_id = 'FBA17QJND724'


def remove_filters_and_search(file_path, target_id):
    try:
        # Load the Excel file with openpyxl engine to remove filters
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            workbook = writer.book
            for sheet in workbook.worksheets:
                # Remove auto-filters if present
                if sheet.auto_filter.ref:
                    sheet.auto_filter.ref = None
            workbook.save(file_path)

        # Load the Excel file with pandas to search for the target ID
        df = pd.read_excel(file_path, engine='openpyxl')

        # Check if the file has a column named "Shipment ID"
        if 'Shipment ID' in df.columns:
            # Search for the target ID in the "Shipment ID" column (Column B)
            if df['Shipment ID'].astype(str).str.contains(target_id).any():
                print(f'ID "{target_id}" found in file: {file_path}')
                return True
    except Exception as e:
        print(f'Error processing file {file_path}: {e}')
    return False


# Loop through all Excel files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        if remove_filters_and_search(file_path, target_id):
            break
else:
    print(f'ID "{target_id}" not found in any files.')
