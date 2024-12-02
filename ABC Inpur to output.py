import pandas as pd
import os
from openpyxl import Workbook

# Define input and output paths
input_folder = r'E:\ABC Input'
output_folder = r'E:\ABC Output File'

# Dynamically find the CSV file in the input folder
input_file = None
for file in os.listdir(input_folder):
    if file.endswith(".csv"):
        input_file = os.path.join(input_folder, file)
        break

if input_file is None:
    raise FileNotFoundError("No CSV file found in the specified input folder.")

# Read the CSV file
df = pd.read_csv(input_file)

# Print the columns in the DataFrame for debugging
print("Columns in the CSV file:", df.columns.tolist())

# Ensure column 'Reference ID' exists
if 'Reference ID' not in df.columns:
    raise KeyError("Column 'Reference ID' does not exist in the CSV file.")

# Apply formatting and remove rows that start with 'FBA' in column 'Reference ID'
df_filtered = df[~df['Reference ID'].astype(str).str.startswith('FBA')]
df_filtered['Reference ID'] = df_filtered['Reference ID'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

# Remove any rows that have no numeric data in column 'Reference ID'
df_filtered = df_filtered[df_filtered['Reference ID'].str.strip().astype(bool)]

# Create an output Excel file with the name "ABC.xlsx"
output_file = os.path.join(output_folder, 'ABC.xlsx')
wb = Workbook()
ws = wb.active
ws.title = "ABC"  # Set the sheet name to "ABC"

# Write the filtered 'Reference ID' data to the Excel file
for r_idx, value in enumerate(df_filtered['Reference ID'], start=1):
    ws.cell(row=r_idx, column=1, value=value)  # Write only 'Reference ID' column

# Set the width of the first column to 30 characters
ws.column_dimensions['A'].width = 30

# Save the workbook
wb.save(output_file)

print(f"Data processed and saved to {output_file}")
