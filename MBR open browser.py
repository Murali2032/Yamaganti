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

# Open the website
driver.get("https://www.vectorworkflow.com/")

# Maximize the browser window
driver.maximize_window()

# Create an instance of WebDriverWait
wait = WebDriverWait(driver, 10)  # Wait up to 10 seconds

# Missing Bill Report

driver.get("https://www.vectorworkflow.com/#/main/garage/missinginvoicereport")
time.sleep(15)

dropdown = driver.find_element(by=By.XPATH, value = "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div")
dropdown.click()
print("Clicked on Dropdown")
time.sleep(15)

stat_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li")))
stat_select.click()
print("Clicked on Active")
time.sleep(15)

search_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
search_select.click()
print("Clicked on Search")
time.sleep(10)

print("clicked")
exl_down_btn = driver.find_element(by=By.XPATH, value =  "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")
print("Found")
exl_down_btn.click()
print("Downloaded")
time.sleep(10)



