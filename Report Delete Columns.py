import pandas as pd
import os

# Define the file path
file_path = r'E:\Report Header'

# Read the file into a DataFrame
try:
    df = pd.read_csv(file_path)  # Assuming it's a CSV file. Change this if it's a different format.
except FileNotFoundError:
    print(f"File not found at {file_path}")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    exit()

# Define columns to remove (0-based index)
columns_to_remove = ['d', 'f', 'h', 'I', 'j', 'n', 'o', 'p', 'q', 'r', 's']

# Check if columns to remove exist in the DataFrame
existing_columns_to_remove = [col for col in columns_to_remove if col in df.columns]

if not existing_columns_to_remove:
    print("None of the specified columns are in the DataFrame.")
else:
    # Drop the specified columns
    df = df.drop(columns=existing_columns_to_remove)

    # Save the modified DataFrame back to a new file
    new_file_path = os.path.join(os.path.dirname(file_path), 'Modified_Report_Header.csv')
    df.to_csv(new_file_path, index=False)
    print(f"Modified file saved to {new_file_path}")
