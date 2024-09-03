import pyperclip
import openpyxl
import pyautogui
import time
import keyboard

# Load the Excel file
workbook = openpyxl.load_workbook('C:/Users/YamagantiMuraliKrish/Desktop/ABC.xlsx')

# Select the desired sheet
sheet = workbook['ABC']

# Extract the data from the Excel sheet
data = []
for row in sheet.iter_rows(values_only=True):
    data.append(str(row[0]))  # Assuming the data is in the first column (A)

# Additional information for each batch
additional_info = [
    "Do 18 months Audit and issue refund for missing FNSKUs",
    "Do issue refund for missing units in FNSKUs",
    "Provide detailed shipment information of FNSKU",
    "Inventory lost in FBA warehouse",
    "Request to reconcile or reimburse missing inventory in fulfillment centers"
]

# Set the starting row and batch size
start_row = 1
batch_size = 10

# Get the total number of rows in column A
total_rows = len(data)

# Iterate through the rows in batches
while start_row <= total_rows:
    # Select the rows in the current batch
    end_row = min(start_row + batch_size - 1, total_rows)
    batch_data = data[start_row - 1:end_row]

    # Join the data with newlines
    data_text = '\n'.join(batch_data)

    # Add additional information for each batch
    additional_info_text = '\n'.join(additional_info)

    # Wait for the user to press the '.' key
    print("Press '.' to paste data and additional information.")
    keyboard.wait('.')

    # Copy the combined data and additional information to the clipboard
    combined_text = f"{data_text}\n\n{additional_info_text}"
    pyperclip.copy(combined_text)

    # Simulate keyboard input to paste in the browser search bar
    pyautogui.hotkey('ctrl', 'v')

    # Wait for the browser to load and refresh
    time.sleep(10)

    # Refresh the browser (customize this based on your browser's refresh shortcut)

    # Update the starting row for the next batch
    start_row += batch_size