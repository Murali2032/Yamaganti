import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # Import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
import pyperclip
import pyautogui

# Specify the debugging address for the already opened Chrome browser
debugger_address = 'localhost:8989'

c_options = Options()
c_options.add_experimental_option("debuggerAddress", debugger_address)

# Initialize the WebDriver with the existing Chrome instance
driver = webdriver.Chrome(options=c_options)

try:
    # Open the page and print the title
    driver.get('https://sellercentral.amazon.com/reportcentral/REIMBURSEMENTS/0/%7B%22filters%22:[%22%22,%22%22,%22%22,%22%22],%22pageOffset%22:1,%22searchDays%22:365%7D')
    print(driver.title)

    # Directory containing the Excel files
    directory_path = r'C:\Users\YamagantiMuraliKrish\Videos\Reimbrusement'

    # List all files in the directory
    files = os.listdir(directory_path)

    # Filter for Excel files
    excel_files = [f for f in files if f.endswith('.xlsx')]

    # Check if there are any Excel files
    if not excel_files:
        raise FileNotFoundError("No Excel files found in the directory.")

    # Optionally, sort the files by modification time or use other criteria
    excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(directory_path, x)), reverse=True)

    # Select the most recent Excel file (or choose based on your criteria)
    latest_excel_file = excel_files[0]
    excel_path = os.path.join(directory_path, latest_excel_file)

    # Load the selected Excel file
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    # Read the first cell's data from the 'reimbursement-id' column
    reimbursement_id = sheet['B2'].value  # Adjust the cell reference based on your actual data location

    # Convert reimbursement_id to string if it's not already
    if reimbursement_id is not None:
        reimbursement_id = str(reimbursement_id)

        time.sleep(5)

        # Click on the input field
        pyautogui.click(x=831, y=550, duration=0.3)
        time.sleep(2)

        # Paste the data into the active element
        pyautogui.typewrite(reimbursement_id)  # Use typewrite to paste the data
        time.sleep(2)
        pyautogui.click(x=412, y=726, duration=0.3)
    else:
        print("Reimbursement ID is None.")

        pyautogui.click(x=557, y=548, duration=0.3)

finally:
    # Close the browser
    driver.quit()