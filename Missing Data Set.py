import pandas as pd

# File path to the Excel or CSV file
file_path = r'E:\Missing Bill Report\Missing Invoice Report_1732192206.xlsx'  # Change the file name accordingly

# Read the file based on extension (Excel or CSV)
if file_path.endswith('.xlsx'):
    df = pd.read_excel(file_path)
elif file_path.endswith('.csv'):
    df = pd.read_csv(file_path)
else:
    raise ValueError("Unsupported file format. Please use .xlsx or .csv.")

# Check the first few rows to see the column names and structure
print(df.head())

# Filter the 'Missing Status' column for 'Missing' and 'Anticipated'
missing_count = df[df['Missing Status'] == 'Missing'].shape[0]
anticipated_count = df[df['Missing Status'] == 'Anticipated'].shape[0]

# Print the counts
print(f"Missing Count: {missing_count}")
print(f"Anticipated Count: {anticipated_count}")
