import pyperclip
import openpyxl
import pyautogui
import keyboard
import time

# Load the Excel file
workbook = openpyxl.load_workbook('S:/ABC.xlsx')

# Select the desired sheet
sheet = workbook['ABC']

# Get all data from columns A and B
column_a_data = [cell.value for cell in sheet['A'] if cell.value]
column_b_data = [cell.value for cell in sheet['B'] if cell.value]

# Convert data to strings (in case they aren't already)
column_a_data = [str(data) for data in column_a_data]
column_b_data = [str(data) for data in column_b_data]

# Ensure there's data in both columns
if not column_a_data or not column_b_data:
    print("No data found in column A or column B.")
    exit()

# Initialize counters for columns A and B
a_index = 0
b_index = 0

print("Press '.' to paste data from Column A. Press '\\' to paste data from Column B. Press 'Esc' to exit.")

# Continuous loop to wait for keyboard input
while True:
    # Wait for the '.' key to be pressed
    if keyboard.is_pressed('.'):
        # Copy the current cell data from Column A to the clipboard
        pyperclip.copy(column_a_data[a_index])

        # Simulate keyboard input to paste in the browser
        pyautogui.hotkey('ctrl', 'v')

        # Move to the next item in Column A (looping back if at the end)
        a_index = (a_index + 1) % len(column_a_data)

        # Delay to avoid repeated triggering
        time.sleep(0.3)

    # Wait for the '\' key to be pressed
    if keyboard.is_pressed('\\'):
        # Copy the current cell data from Column B to the clipboard
        pyperclip.copy(column_b_data[b_index])

        # Simulate keyboard input to paste in the browser
        pyautogui.hotkey('ctrl', 'v')

        # Move to the next item in Column B (looping back if at the end)
        b_index = (b_index + 1) % len(column_b_data)

        # Delay to avoid repeated triggering
        time.sleep(0.3)

    # Exit the loop if the 'Esc' key is pressed
    if keyboard.is_pressed('esc'):
        print("Exiting script.")
        break
