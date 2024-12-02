# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 16:39:32 2024
@author: JohnK
"""
from email.mime.multipart import MIMEMultipart
import pyodbc
import glob
import pandas as pd
import os
import shutil
import time
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import encoders
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import smtplib

#Selenium Code
# Define paths for the reports
account_report_dir = "E:\\List Report"
missing_bill_report_dir = "E:\\Missing Bill Report"
temp_download_dir = "E:\\TempDownload"

def clear_directory(folder_path):
    """Delete all files in the given folder."""
    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                    print(f"Deleted file: {file_path}")
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    print(f"Deleted folder: {file_path}")
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

def move_downloaded_files(destination_folder):
    """Move files from the temp download directory to the specified folder."""
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(temp_download_dir):
        source_path = os.path.join(temp_download_dir, filename)
        dest_path = os.path.join(destination_folder, filename)
        shutil.move(source_path, dest_path)
        print(f"Moved {filename} to {destination_folder}")

# Clear existing files in both target folders
clear_directory(account_report_dir)
clear_directory(missing_bill_report_dir)
clear_directory(temp_download_dir)

try:
    # Set up Chrome options with download preferences
    chrome_options = Options()
    prefs = {
        "download.default_directory": temp_download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('window-size=1920x1080')
    # chrome_options.add_argument('--headless')  # Uncomment for headless mode if needed

    # Initialize WebDriver using webdriver-manager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # Login to the website
    driver.get("https://www.vectorworkflow.com/")
    driver.maximize_window()

    userid = 'muralik'
    user_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[2]/div/input")))
    user_input.send_keys(userid)

    password = 'welcome'
    pass_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[3]/div/input")))
    pass_input.send_keys(password)

    sign_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-login/div/div[1]/div/div[2]/div/div[2]/div[5]/button")))
    sign_btn.click()
    time.sleep(5)

    # Download Account List Report
    driver.get("https://www.vectorworkflow.com/#/main/garage/AccountListReport")
    time.sleep(10)

    dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[2]/span")))
    dropdown.click()
    time.sleep(2)

    active_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
    active_option.click()
    time.sleep(2)

    search_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[2]/button[1]/span")))
    search_btn.click()
    time.sleep(5)

    download_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
    download_btn.click()
    time.sleep(15)
    print("Account List Report Downloaded")

    # Move Account List Report to the target folder
    move_downloaded_files(account_report_dir)

    # Download Missing Bill Report
    driver.get("https://www.vectorworkflow.com/#/main/garage/missinginvoicereport")
    time.sleep(10)

    dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div")))
    dropdown.click()
    time.sleep(2)

    active_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li")))
    active_option.click()
    time.sleep(2)

    search_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
    search_btn.click()
    time.sleep(5)

    download_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
    download_btn.click()
    time.sleep(10)
    print("Missing Bill Report Downloaded")

    # Move Missing Bill Report to the target folder
    move_downloaded_files(missing_bill_report_dir)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    driver.quit()

#Selenium Ends

# Function to get the latest Excel file from a directory
def get_latest_file_from_dir(directory):
    # Search for all Excel files in the specified directory
    excel_files = glob.glob(os.path.join(directory, "*.xlsx"))

    # Check if there are any Excel files in the directory
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in {directory}")

    # Get the latest Excel file based on modification time
    latest_file = max(excel_files, key=os.path.getmtime)
    return latest_file

# Define the directories for the reports
account_list_directory = "E:\\List Report"
missing_invoice_directory = "E:\\Missing Bill Report"

# Get the latest files from the directories
try:
    account_list_report_path = get_latest_file_from_dir(account_list_directory)
    missing_invoices_report_path = get_latest_file_from_dir(missing_invoice_directory)

    # Read the Excel files dynamically
    Account_List_Report = pd.read_excel(account_list_report_path)
    Missing_Invoices_Report = pd.read_excel(missing_invoices_report_path)
except FileNotFoundError as e:
    print(e)
    exit()

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
select BD.ImageFileName, I.InvoiceNumber, I.InvoiceStatus, isnull(convert(varchar, I.ServiceStartDate, 101),'None') as ServiceStartDate,
isnull(convert(varchar,I.Invoicedate,101), 'None') as InvoiceDate,
isnull(convert(varchar, I.ServiceEnddate, 101), 'None') as ServiceEndDate, isnull(I.CurrentCharges, 0) as CurrentCharges,
isnull(E.ExceptionDescription, 'Blank') ExceptionDescription,
A.AccountNumber, P.propertyName, C.ClientName, CT.ContractNo, V.VendorName
from Invoice I                                    
inner join BatchDetails BD on BD.BatchDetailId = I.BatchDetailId                                    
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

# Add new columns to Account_List_Report
New_Columns_Acct_List = pd.DataFrame({
    'Age': [None] * len(Account_List_Report),
    'Today': [datetime.now().date()] * len(Account_List_Report),
    'LastInvoiceDate': [None] * len(Account_List_Report),
    'Missing': [None] * len(Account_List_Report)
})
Account_List_Report = pd.concat([New_Columns_Acct_List, Account_List_Report], axis=1)

if 'Account #' in Account_List_Report.columns:
    Account_List_Report['Missing'] = ~Account_List_Report['Account #'].isin(Missing_Invoices_Report['AccountNumber'])
else:
    print("Column 'Account #' not found in 'Account List Report'.")

# Filter for missing accounts
Account_List_Report = Account_List_Report[Account_List_Report['Missing'] == True]

# Prepare Open_Invoices_Report for merging
Open_Invoices_Report['InvoiceDate'] = pd.to_datetime(Open_Invoices_Report['InvoiceDate'], format='%m/%d/%Y', errors='coerce')
Open_Invoices_Report = Open_Invoices_Report.dropna(subset=['InvoiceDate']).sort_values(['AccountNumber', 'InvoiceDate'], ascending=[True, False])
Open_Invoices_Report = Open_Invoices_Report.drop_duplicates(subset='AccountNumber', keep='first').reset_index(drop=True)

if 'Account #' in Account_List_Report.columns and 'AccountNumber' in Open_Invoices_Report.columns:
    Account_List_Report = pd.merge(
        Account_List_Report,
        Open_Invoices_Report[['AccountNumber', 'InvoiceDate']],
        left_on='Account #',
        right_on='AccountNumber',
        how='left'
    )
    Account_List_Report['LastInvoiceDate'] = Account_List_Report['InvoiceDate']
    Account_List_Report = Account_List_Report.drop(columns=['AccountNumber', 'InvoiceDate'])

# Convert LastInvoiceDate to datetime and handle errors
try:
    Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], errors='coerce')
except Exception as e:
    print(f"Error converting to datetime: {e}")
    sample_date = Account_List_Report['LastInvoiceDate'].iloc[0]
    if '/' in str(sample_date):
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], format='%m/%d/%Y', errors='coerce')
    elif '-' in str(sample_date):
        Account_List_Report['LastInvoiceDate'] = pd.to_datetime(Account_List_Report['LastInvoiceDate'], format='%Y-%m-%d', errors='coerce')
    else:
        print("Unable to determine date format. Please provide the correct format.")

# Filter by valid years
valid_years = [2022, 2023, 2024]
Account_List_Report['Year'] = Account_List_Report['LastInvoiceDate'].dt.year
Account_List_Report = Account_List_Report[Account_List_Report['Year'].isin(valid_years)]

# Convert Today to datetime for Age calculation
Account_List_Report['Today'] = pd.to_datetime(Account_List_Report['Today'])
Account_List_Report['Age'] = (Account_List_Report['Today'] - Account_List_Report['LastInvoiceDate']).dt.days

# Exclude rows where Age is between 0 and 29 days
Account_List_Report = Account_List_Report[(Account_List_Report['Age'] < 0) | (Account_List_Report['Age'] > 29)]

# Delete all files in the output folder
output_folder = "C:\\Users\\YamagantiMuraliKrish\\Downloads\\Account_List_Report"
files = glob.glob(os.path.join(output_folder, "*"))  # Delete all files regardless of format

for file in files:
    try:
        os.remove(file)
        print(f"Deleted file: {file}")
    except Exception as e:
        print(f"Error deleting file {file}: {e}")

# Save to Excel file
output_path = os.path.join(output_folder, "Hello123.xlsx")
Account_List_Report.to_excel(output_path, index=False)
print(f"Output saved to {output_path}")

# SharePoint credentials
sharepoint_site_url = "https://rsllc.sharepoint.com/sites/DataPulse/"
username = "reports@vector97.com"
password = "bylsrbnxybdblkkw"

# Local file path and target SharePoint folder path
sharepoint_folder_url = "/sites/DataPulse/Shared Documents/Revenue Generation Dashboard"

try:
    # Authenticate and connect to SharePoint
    ctx_auth = AuthenticationContext(sharepoint_site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(sharepoint_site_url, ctx_auth)
        # Try to upload using folder.upload_file
        with open('C:\\Users\\YamagantiMuraliKrish\\Downloads\\Account_List_Report', 'rb') as file_content:
            file_name = os.path.basename("C:\\Users\\YamagantiMuraliKrish\\Downloads\\Account_List_Report")
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder_url)
            target_folder.upload_file(file_name, file_content).execute_query()
            print(f"Uploaded {file_name} to SharePoint successfully.")
    else:
        print("Authentication failed.")
except Exception as e:
    print(f"An error occurred: {str(e)}")

##Start Email Funtionality

fromaddr = "muralik@vector97.com"
toaddr = ['muralik@vector97.com']
cc = ['muralik@vector97.com']
bcc = ['johnk@vector97.com', ]


msg = MIMEMultipart()

msg["From"] = fromaddr
msg["To"] = ", ".join(toaddr)
msg["Cc"] = ", ".join(cc)
msg['Bcc'] = ', '.join(bcc)
msg["Subject"] = (
        "Variable Missing Bill Report | " + date_time_file + " | " + time_T + " CST "
)

body = (
        "Hello Team,<br><br>Please find the attached Variable Missing Bill Report"
        + " @ "
        + date_time_file
        + " | "
        + time_T
        + " CST "
        + "<br><br>"
        + "Thanks,"
        + "<br>"
        + "<b>Team Vector97</b>"
)

msg.attach(MIMEText(body, "html"))

filename = '"Variable Missing Bill Report ' + date_time_file + time_string + ".xlsx"
attachment = open('C:\\Users\\YamagantiMuraliKrish\\Downloads\\Account_List_Report', "rb")

part = MIMEBase("application", "octet-stream")
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header("Content-Disposition", "attachment; filename= %s" % filename)

msg.attach(part)

server = smtplib.SMTP("smtp.office365.com", 587)
server.starttls()
cwd = os.getcwd()
f = open(cwd + "\\" + "Password.txt", "r")
password = f.read()
f.close()
server.login(fromaddr, password)
text = msg.as_string()
server.sendmail(fromaddr, (toaddr + cc + bcc), text)
server.quit()


