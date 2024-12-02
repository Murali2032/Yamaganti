import time
import openpyxl
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Specify the path to the ChromeDriver executable
chrome_path = "D:\\Chrome Driver\\chromedriver.exe"

# Initialize the WebDriver with the existing Chrome instance
service = Service(chrome_path)
opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
driver = webdriver.Chrome(service=service, options=opt)

# Wait object for handling waits
wait = WebDriverWait(driver, 20)

# Login to the website
driver.get("https://www.vectorworkflow.com/")
driver.maximize_window()

userid = 'muralik'
user_input = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[2]/div/input")))
user_input.click()
time.sleep(10)
user_input.send_keys(userid)
time.sleep(10)

password = 'welcome'
pass_input = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[3]/div/input")))
pass_input.click()
time.sleep(5)
pass_input.send_keys(password)

sign_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[5]/button")))
sign_btn.click()
time.sleep(5)

# Account List Report
driver.get("https://www.vectorworkflow.com/#/main/garage/AccountListReport")
print(driver.title)
time.sleep(10)

# Wait for dropdown to be clickable and select it
dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[2]/span")))
dropdown.click()
print("Clicked on Dropdown")
time.sleep(5)

dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
dropdown.click()
print("Clicked on Active")
time.sleep(5)

Search = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[2]/button[1]/span")))
Search.click()
print("Clicked on Search")
time.sleep(5)

# Download Button - Wait until clickable and click
Download = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
Download.click()
time.sleep(15)
print("Report Downloaded")

# Missing Bill Report
driver.get("https://www.vectorworkflow.com/#/main/garage/missinginvoicereport")
time.sleep(15)

dropdown = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div")
dropdown.click()
print("Clicked on Dropdown")
time.sleep(10)

stat_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li")))
stat_select.click()
print("Clicked on Active")
time.sleep(20)

search_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
search_select.click()
print("Clicked on Search")
time.sleep(10)

exl_down_btn = driver.find_element(by=By.XPATH, value="/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")
print("Found")
exl_down_btn.click()
print("Downloaded")
time.sleep(10)

# Handle downloaded files (Optional)