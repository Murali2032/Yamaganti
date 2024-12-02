import os
import shutil
import time
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

# Define paths for the reports
account_report_dir = "E:\\List Report"
missing_bill_report_dir = "E:\\Missing Bill Report"
temp_download_dir = "E:\\TempDownload"
sharepoint_site_url = "https://rsllc.sharepoint.com/sites/DataPulse/"
sharepoint_folder_url = "/sites/DataPulse/Shared Documents/Variable Missing Bill Report"
username = "reports@vector97.com"
password = "bylsrbnxybdblkkw"


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
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[2]/span")))
    dropdown.click()
    time.sleep(2)

    active_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[1]/div/div[10]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li/span[1]")))
    active_option.click()
    time.sleep(2)

    search_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/report-searchfields/p-panel/div/div[2]/div/div[2]/button[1]/span")))
    search_btn.click()
    time.sleep(5)

    download_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/account-list-report/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
    download_btn.click()
    time.sleep(15)
    print("Account List Report Downloaded")

    # Move Account List Report to the target folder
    move_downloaded_files(account_report_dir)

    # Download Missing Bill Report
    driver.get("https://www.vectorworkflow.com/#/main/garage/missinginvoicereport")
    time.sleep(10)

    dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div")))
    dropdown.click()
    time.sleep(2)

    active_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[1]/div[7]/span/ddl-accountstatus/span/p-dropdown/div/div[3]/div/ul/p-dropdownitem[1]/li")))
    active_option.click()
    time.sleep(2)

    search_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-panel/div/div[2]/div/div[3]/button[1]")))
    search_btn.click()
    time.sleep(5)

    download_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH,
         "/html/body/app-root/app-main/div/div/div[1]/garage-master/div/missing-invoice-report/div/p-tabview/div/div/p-tabpanel[1]/div/app-vectorgrid/p-table/div/div[1]/button[4]/span")))
    download_btn.click()
    time.sleep(10)
    print("Missing Bill Report Downloaded")

    # Move Missing Bill Report to the target folder
    move_downloaded_files(missing_bill_report_dir)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    driver.quit()

# Define the downloaded file path for upload
new_file_path = os.path.join(account_report_dir, "Account_List_Report.xlsx")  # Example filename
# If the file is in missing_bill_report_dir
# new_file_path = os.path.join(missing_bill_report_dir, "Missing_Bill_Report.xlsx")

try:
    # Authenticate and connect to SharePoint
    ctx_auth = AuthenticationContext(sharepoint_site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(sharepoint_site_url, ctx_auth)

        # Try to upload using folder.upload_file
        with open(new_file_path, 'rb') as file_content:
            file_name = os.path.basename(new_file_path)
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder_url)
            target_folder.upload_file(file_name, file_content).execute_query()
            print(f"Uploaded {file_name} to SharePoint successfully.")
    else:
        print("Authentication failed.")
except Exception as e:
    print(f"An error occurred: {str(e)}")

# Email functionality
fromaddr = "reports@vector97.com"
toaddr = ['muralik@vector97.com']
cc = ['muralik@vector97.com']
bcc = ['johnk@vector97.com']

msg = MIMEMultipart()
msg["From"] = fromaddr
msg["To"] = ", ".join(toaddr)
msg["Cc"] = ", ".join(cc)
msg['Bcc'] = ', '.join(bcc)
msg["Subject"] = (
        "Variable Missing Bill Report | " + time.strftime("%Y-%m-%d") + " | " + time.strftime("%H:%M:%S") + " CST "
)

body = (
    "Hello Team, Please find the attached missing bill report. "
    "Kindly take necessary action.\n\nThank you,\nTeam"
)

msg.attach(MIMEText(body, 'plain'))

# Attach the files
attachment1 = open(new_file_path, "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment1.read())
encoders.encode_base64(part)
part.add_header(
    'Content-Disposition', f"attachment; filename={os.path.basename(new_file_path)}"
)
msg.attach(part)

# Send email
try:
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(fromaddr, password)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr + cc + bcc, text)
    print("Email sent successfully")
except Exception as e:
    print(f"Failed to send email: {str(e)}")
finally:
    server.quit()