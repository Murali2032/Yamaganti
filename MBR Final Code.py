import os
import shutil
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # Make sure this is imported
from selenium.webdriver.support import expected_conditions as EC

try:
    # Specify the folder path  *** This is to delete existing files in folder ***
    folder_path = "E:\\Report"

    # Check if the folder exists
    if os.path.exists(folder_path):
        # Iterate through all the files and subfolders in the folder
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                # Check if it is a file or directory
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Remove the file or link
                    print(f"Deleted file: {file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # Remove the directory
                    print(f"Deleted folder: {file_path}")
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
    else:
        print(f"The folder {folder_path} does not exist.")
except Exception as e:
    print(f"An error occurred while deleting files: {e}")

try:
    # Define the path to chromedriver.exe
    driver_path = "C:\\Users\\YamagantiMuraliKrish\\Yamaganti\\chromedriver.exe"

    # Directory to save the report
    download_dir = "E:\\Report"

    # Ensure the directory exists
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # Set up Chrome options (including download preferences)
    chrome_options = webdriver.ChromeOptions()

    # Set download preferences
    prefs = {
        "download.default_directory": download_dir,  # Set the default download directory
        "download.prompt_for_download": False,       # Disable the download prompt
        "directory_upgrade": True,                   # Automatically replace existing files
        "safebrowsing.enabled": True                 # Enable safe browsing
    }

    # Apply the preferences to Chrome options
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.add_argument('window-size=1920x1080')  # Optional: Set the browser window size
    # chrome_options.add_argument('--headless')  # Uncomment for headless mode if needed

    # Initialize WebDriver with the specified driver path and options
    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)

    # Login to the website
    driver.get("https://www.vectorworkflow.com/")
    driver.maximize_window()

    userid = 'muralik'
    user_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[2]/div/input")))

    user_input.click()
    time.sleep(10)
    user_input.send_keys(userid)
    time.sleep(10)

    password = 'welcome'
    pass_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[3]/div/input")))

    pass_input.click()
    time.sleep(5)
    pass_input.send_keys(password)

    sign_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[5]/button")))
    sign_btn.click()
    time.sleep(5)

    # Account List Report
    driver.get("https://www.vectorworkflow.com/#/main/garage/AccountListReport")
    print(driver.title)
    time.sleep(10)

    # Wait for dropdown to be clickable and select it
    dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[2]/span")))
    dropdown.click()
    print("Clicked on Dropdown")
    time.sleep(5)

    dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
    dropdown.click()
    print("Clicked on Active")
    time.sleep(5)

    Search = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[2]/button[1]/span")))
    Search.click()
    print("Clicked on Search")
    time.sleep(5)

    # Download Button - Wait until clickable and click
    Download = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
    Download.click()
    time.sleep(15)
    print("Report Downloaded")

    # Missing Bill Report
    driver.get("https://www.vectorworkflow.com/#/main/garage/missinginvoicereport")
    time.sleep(15)

    dropdown = driver.find_element(by=By.XPATH,value="/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div")
    dropdown.click()
    print("Clicked on Dropdown")
    time.sleep(10)

    stat_select = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li")))
    stat_select.click()
    print("Clicked on Active")
    time.sleep(20)

    search_select = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
    search_select.click()
    print("Clicked on Search")
    time.sleep(10)

    exl_down_btn = driver.find_element(by=By.XPATH,value="/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")
    print("Found")
    exl_down_btn.click()
    print("Downloaded")
    time.sleep(10)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    driver.quit()  # Ensure the browser is closed at the end
