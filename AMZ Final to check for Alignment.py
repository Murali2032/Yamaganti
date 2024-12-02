import os
import pandas as pd

# Path to the folder containing Excel files (including subfolders)
root_directory = r"C:\Users\YamagantiMuraliKrish\Music\November"

# List of required columns
required_columns = [
    "Date", "Reimbursement ID", "Claims Number", "Reimbursement Type",
    "FNSKU", "Currency", "Reimbursement Total"
]

# List of valid currencies
valid_currencies = ["USD", "INR", "CAD", "JPY", "EUR", "AUD", "GBP"]


def check_excel_file(file_path):
    """
    Check if the Excel file contains the required columns and data conditions.
    Return the missing details if the file is invalid.
    """
    missing_columns = []
    invalid_currencies = False
    invalid_reimbursement_type = False
    invalid_reimbursement_total = False

    try:
        # Read the Excel file
        data = pd.read_excel(file_path)

        # Check for missing columns
        missing_columns = [col for col in required_columns if col not in data.columns]

        # Check if "Currency" column contains valid currencies
        currencies_in_file = set(data["Currency"].dropna().unique())
        if not any(currency in valid_currencies for currency in currencies_in_file):
            invalid_currencies = True  # No valid currencies found

        # Check if "Reimbursement Type" contains any 0 values
        if any(data["Reimbursement Type"] == 0):
            invalid_reimbursement_type = True  # Contains 0 in "Reimbursement Type"

        # Check if "Reimbursement Total" contains any 0 values
        if any(data["Reimbursement Total"] == 0):
            invalid_reimbursement_total = True  # Contains 0 in "Reimbursement Total"

        # If there are any issues, return them
        if missing_columns or invalid_currencies or invalid_reimbursement_type or invalid_reimbursement_total:
            return False, missing_columns, invalid_currencies, invalid_reimbursement_type, invalid_reimbursement_total

        return True, None, None, None, None  # All checks passed
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return False, None, None, None, None


def process_files(directory):
    """
    Process all Excel files in the given directory and its subdirectories.
    Check if they contain the required columns and data, and print issues for invalid files.
    """
    valid_files = []  # List to store valid file names
    invalid_files = []  # List to store invalid file names

    # Traverse all files in the directory and subdirectories using os.walk
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(root, filename)

                # Check the Excel file
                is_valid, missing_columns, invalid_currencies, invalid_reimbursement_type, invalid_reimbursement_total = check_excel_file(
                    file_path)

                if is_valid:
                    valid_files.append(file_path)
                else:
                    invalid_files.append(file_path)
                    # Print the issues for invalid files
                    print(f"\nIssues in file: {file_path}")
                    if missing_columns:
                        print(f"  Missing columns: {', '.join(missing_columns)}")
                    if invalid_currencies:
                        print("  Invalid or missing currencies in the 'Currency' column.")
                    if invalid_reimbursement_type:
                        print("  'Reimbursement Type' contains 0 values.")
                    if invalid_reimbursement_total:
                        print("  'Reimbursement Total' contains 0 values.")

    # Print out the files that passed the checks (valid files)
    if valid_files:
        print("\nThe following files are valid:")
        for file in valid_files:
            print(file)
    else:
        print("\nNo valid files found.")

    # Print out the files that failed the checks (invalid files)
    if invalid_files:
        print("\nThe following files are invalid:")
        for file in invalid_files:
            print(file)
    else:
        print("\nNo invalid files found.")


if __name__ == "__main__":
    print("Starting to process Excel files...")
    process_files(root_directory)
    print("Processing completed!")
