# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 12:12:35 2023

@author: Murali
"""

import openpyxl
import pyperclip
import pyautogui
import time

# Load the Excel file
workbook = openpyxl.load_workbook('C:\\Users\\Murali\\OneDrive - Refuse Specialists LLC\\Desktop\\Sheet.xlsx')

# Select the desired sheet
sheet = workbook['Sheet']

# Extract the data from the Excel sheet
data = []
for row in sheet.iter_rows(values_only=True):
    data.append(str(row[0]))  # Assuming the data is in the first column (A)

# Set the starting row and batch size
start_row = 1
batch_size = 1

# Get the total number of rows in column A
total_rows = len(data)

# Iterate through the rows in batches
while start_row <= total_rows:
    # Select the rows in the current batch
    end_row = min(start_row + batch_size - 1, total_rows)
    batch_data = data[start_row - 1:end_row]

    # Join the data with newlines
    data_text = '\n'.join(batch_data)

    # Copy the data to the clipboard
    pyperclip.copy(data_text)

    # Simulate keyboard input to paste in the browser search bar
    pyautogui.hotkey('ctrl', 'v')

    # Wait for the browser to load and refresh
    time.sleep(7)

    # Refresh the browser (customize this based on your browser's refresh shortcut)

    # Update the starting row for the next batch
    start_row += batch_size