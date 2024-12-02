import time
import openpyxl
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pyautogui

# Specify the path to the ChromeDriver executable
chrome_path = "D:\\Chrome Driver\\chromedriver.exe"

# Initialize the WebDriver with the existing Chrome instance
service = Service(chrome_path)
opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
driver = webdriver.Chrome(service=service, options=opt)

try:
    # Load the Excel file
    workbook = openpyxl.load_workbook('S:\\ABC.xlsx')

    # Select the desired sheet
    sheet = workbook['ABC']

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

        # Run the code block
        driver.get("https://sellercentral.amazon.com/help/hub/reference/GGEV4254LJJ9BAEG")
        print(driver.title)

        # Get the initial height of the page
        initial_height = driver.execute_script("return document.body.scrollHeight")
        time.sleep(3)

        # Scroll down continuously until reaching the end of the page
        while True:
            # Scroll down by a certain amount
            pyautogui.scroll(-1000)  # Adjust the scrolling amount as needed

            # Wait for the page to load and adjust scrolling speed if needed
            time.sleep(1)  # Adjust the sleep time as needed

            # Check if the page height has changed
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == initial_height:
                # Stop scrolling if the page height remains the same
                break
            else:
                # Update the initial height
                initial_height = new_height

        # Click on the input field
        pyautogui.click(x=707, y=566, duration=0.3)
        time.sleep(2)

        # Paste the data into the active element
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)

        # Print the data that has been pasted into the input field
        active_element = driver.switch_to.active_element
        input_value = active_element.get_attribute('value')
        print("Data pasted into the input field:", input_value)

        # Click on the submit button
        pyautogui.click(x=888, y=619, duration=0.3)
        time.sleep(3)

        pyautogui.click(x=874, y=708, duration=0.3)
        time.sleep(3)

        # Create case
        pyautogui.click(x=888, y=619, duration=0.3)
        time.sleep(3)

        # Create case
        pyautogui.click(x=871, y=618, duration=0.3)
        time.sleep(3)

        # Take screenshot and save it with a dynamically generated filename
        screenshot_path = f"C:\\Users\\YamagantiMuraliKrish\\Music\\Screenshot\\screenshot_batch_{start_row}.png"
        driver.save_screenshot(screenshot_path)

        # Check if a certain element is visible then click or run programmed
        # Define the coordinates where you want to check for visibility
        x = 741
        y = 677

        def is_area_below_visible(x, y):
            # Define the coordinates of the area below
            area_below_x = x
            area_below_y = y + 50
            return True


        # Check if the area below the specified coordinates is visible
        if is_area_below_visible(x, y):
            # If the area is visible, perform the click action
            pyautogui.click(x=x, y=y, duration=0.3)
            time.sleep(3)

        # Store the handle of the primary window
        primary_window = driver.current_window_handle

        # Switch back to the primary window
        driver.switch_to.window(primary_window)

        start_row += batch_size

finally:
    # Close the WebDriver
    driver.quit()