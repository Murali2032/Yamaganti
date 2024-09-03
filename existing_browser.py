# # chrome.exe --remote-debugging-port=8989 --user-data-dir=D:\Chrome Driver\chromedriver

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

# Set up ChromeOptions and connect to the existing browser
c_options = Options()
c_options.add_experimental_option("debuggerAddress", "localhost:8989")

# Initialize the WebDriver with the existing Chrome instance
driver = webdriver.Chrome(options=c_options)

# Now, you can interact with the already opened Chrome browser
driver.get('https://sellercentral.amazon.com/help/center?mons_redirect=stck_reroute')
print(driver.title)

driver.get("https://sellercentral.amazon.com/help/center")
print(driver.title)

driver.implicitly_wait(10) # seconds

# Find the host element


# You can now interact with elementInsideShadow as needed









