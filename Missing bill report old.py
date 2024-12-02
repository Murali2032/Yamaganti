from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import pyodbc
import pandas as pd
from datetime import datetime
import os

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

# Define the path to chromedriver.exe
driver_path = "C:\\Users\\YamagantiMuraliKrish\\Yamaganti\\chromedriver.exe"

# Set up Chrome options (optional)
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('window-size=1920x1080')  # Optional: Set the browser window size
# chrome_options.add_argument('--headless')  # Uncomment for headless mode if needed

# Initialize WebDriver with the specified driver path
driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)

# Directory to save the report
download_dir = "E:\\Report"

# Ensure the directory exists
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# Specify download preferences
prefs = {
    "download.default_directory": download_dir,  # Set the default download directory
    "download.prompt_for_download": False,       # Disable the download prompt
    "directory_upgrade": True,                   # Automatically replace existing files
    "safebrowsing.enabled": True                 # Enable safe browsing
}

# Open the website
driver.get("https://www.vectorworkflow.com/")

# Maximize the browser window
driver.maximize_window()

# Create an instance of WebDriverWait
wait = WebDriverWait(driver, 10)  # Wait up to 10 seconds

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
time.sleep(10)

# Account List Report
driver.get("https://www.vectorworkflow.com/#/main/garage/AccountListReport")
print('Opened website')
time.sleep(20)

# dropdown = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
# dropdown.click()
# print("Clicked on Account list report")
# time.sleep(10)

Active = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
Active.click()
print("Active Accounts Selected")
time.sleep(10)

Search = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[2]/button[1]/span")))
Search.click()
print("Clicked on Search")
time.sleep(10)

Download = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
Download.click()
time.sleep(15)
print("Report Downloaded")

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

srch_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
srch_select.click()
print("Clicked on Search")
time.sleep(10)

print("clicked")
exl_down_btn = driver.find_element(by=By.XPATH, value =  "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")
print("Found")
exl_down_btn.click()
print("Downloaded")
time.sleep(10)


# Read Excel files
Account_List_Report = pd.read_excel("C:\\Users\\JohnK\\Downloads\\Account List Report_1727120848.xlsx")
Missing_Invoices_Report = pd.read_excel("C:\\Users\\JohnK\\Downloads\\Missing Invoice Report_1727120873.xlsx")

# Database connection
server = "wmssql.database.windows.net"
database = "Vectorproddb"
username = "John"
password = "Quadrant@12$"
driver = "{SQL Server}"
cnxn = pyodbc.connect(
    f"DRIVER={driver};SERVER=tcp:{server};PORT=1433;DATABASE={database};UID={username};PWD={password}"
)

# Execute SQL query
OpenInvoicesReport_query = """SET NOCOUNT ON;
select BD.ImageFileName, I.InvoiceNumber, I.InvoiceStatus, isnull(convert(varchar, I.ServiceStartDate, 101),'None') as ServiceStartDate
, isnull(convert(varchar,I.Invoicedate,101), 'None') as InvoiceDate
, isnull(convert(varchar, I.ServiceEnddate, 101), 'None') as ServiceEndDate, isnull(I.CurrentCharges, 0) as CurrentCharges,
isnull(E.ExceptionDescription, 'Blank') ExceptionDescription,
A.AccountNumber, P.propertyName, C.ClientName, CT.ContractNo, V.VendorName
from Invoice I                                    
    inner join BatchDetails BD on BD.BatchDetailId =I.BatchDetailId                                    
    inner join Batch B on B.BatchId = BD.BatchId                                    
    inner join Accounts A on A.AccountId = I.AccountId                                    
    inner join Contracts CT on CT.ContractId = A.ContractId                                    
    inner join Properties P on P.PropertyId = A.PropertyId                                    
    inner join Clients C on C.clientid = p.Clientid     
    inner join Vendors V on V.VendorId = A.VendorId           
    left join Exceptions E on E.BatchDetailId = I.BatchDetailId    
where 1=1 
and (I.InvoiceStatus not in ('Audited') and I.CreatedDate >= '01/01/2021')
or (I.InvoiceStatus in ('Audited') and I.CreatedDate >= '06/01/2021')
"""
Open_Invoices_Report = pd.read_sql(OpenInvoicesReport_query, cnxn)
cnxn.close()

New_Columns_Acct_List = pd.DataFrame({
    'Age': [None] * len(Account_List_Report),
    'Today': [datetime.now().date()] * len(Account_List_Report),
    'LastInvoiceDate': [None] * len(Account_List_Report),
    'Missing': [None] * len(Account_List_Report)
})

# Concatenate new columns with Account_List_Report
Account_List_Report = pd.concat([New_Columns_Acct_List, Account_List_Report], axis=1)

if 'Account #' in Account_List_Report.columns:
    # Perform the lookup
    Account_List_Report['Missing'] = Account_List_Report['Account #'].isin(Missing_Invoices_Report['AccountNumber'])
    # Invert the boolean to reflect 'Missing' accounts as True
    Account_List_Report['Missing'] = ~Account_List_Report['Missing']

else:
    print("Column 'Account #' not found in 'Account List Report'.")

Account_List_Report = Account_List_Report[Account_List_Report['Missing'] == True]

Open_Invoices_Report = Open_Invoices_Report[['AccountNumber', 'InvoiceDate']]

Open_Invoices_Report['InvoiceDate'] = pd.to_datetime(Open_Invoices_Report['InvoiceDate'],
                                                     format='%m/%d/%Y',
                                                     errors='coerce')
Open_Invoices_Report = Open_Invoices_Report.dropna(subset=['InvoiceDate'])

Open_Invoices_Report = Open_Invoices_Report.sort_values(['AccountNumber', 'InvoiceDate'], ascending=[True, False])

Open_Invoices_Report = Open_Invoices_Report.drop_duplicates(subset='AccountNumber', keep='first')

Open_Invoices_Report = Open_Invoices_Report.reset_index(drop=True)

if 'Account #' in Account_List_Report.columns and 'AccountNumber' in Open_Invoices_Report.columns:
    # Sort Open_Invoices_Report by InvoiceDate and keep the latest date for each AccountNumber
    # latest_invoices = Open_Invoices_Report.sort_values('InvoiceDate').groupby('AccountNumber').last().reset_index()
    # Perform the merge, only bringing in the InvoiceDate column
    Account_List_Report = pd.merge(
        Account_List_Report,
        Open_Invoices_Report[['AccountNumber', 'InvoiceDate']],
        left_on='Account #',
        right_on='AccountNumber',
        how='left'
    )
    Account_List_Report['LastInvoiceDate'] = Account_List_Report['InvoiceDate']
    Account_List_Report = Account_List_Report.drop(columns=['AccountNumber', 'InvoiceDate'])

Account_List_Report.to_excel('C:/Users/JohnK/Downloads/Hello123.xlsx')

Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], errors='coerce')
Account_List_Report['LastInvoiceDate'] = Account_List_Report['LastInvoiceDate'].dt.strftime('%Y-%m-%d')

# Now filter the dates

try:
    Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], errors='coerce')
except Exception as e:
    print(f"Error converting to datetime: {e}")
    print("Attempting alternative conversion methods...")

    # If conversion fails, try to determine the date format
    sample_date = Account_List_Report['LastInvoiceDate'].iloc[0]
    print(f"Sample date: {sample_date}")

    if '/' in str(sample_date):
        # Assuming MM/DD/YYYY format
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'],
                                                                format='%m/%d/%Y', errors='coerce')
    elif '-' in str(sample_date):
        # Assuming YYYY-MM-DD format
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'],
                                                                format='%Y-%m-%d', errors='coerce')
    else:
        print("Unable to determine date format. Please provide the correct format.")

valid_years = [2022, 2023, 2024]
Account_List_Report['Year'] = Account_List_Report['LastInvoiceDate'].dt.year
Account_List_Report = Account_List_Report[Account_List_Report['Year'].isin(valid_years)]

Account_List_Report['Today'] = pd.to_datetime(Account_List_Report['Today'])
Account_List_Report['Age'] = (Account_List_Report['Today'] - Account_List_Report['LastInvoiceDate']).dt.days
Account_List_Report = Account_List_Report[~((Account_List_Report['Age'] >= 0) & (Account_List_Report['Age'] <= 29))]