import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Navigate to the initial URL
driver.get("https://andaze.sharepoint.com/:f:/s/recruit/Eg9lvf5F0VZBmFTBm15jv1gBrXjAw0X33ueT4-B6rIBNSw?e=vuzJbS")
driver.find_element(By.ID, 'txtPassword').send_keys('uDtauXBmbUXe1ZEf')
driver.find_element(By.ID, 'btnSubmitPassword').click()
input("Press Enter to close the browser...")

driver.find_element(By.XPATH, '//button[normalize-space()='Take Home Assignment 2']').click()


driver.implicitly_wait(10)

    # Wait for user input before exiting
input("Press Enter to close the browser...")



# Wait for user input before exiting
#input("Press Enter to close the browser...")




# driver.find_element(By.ID, 'ap_email').send_keys('amazonclaims2@gmail.com')
# driver.find_element(By.ID, 'ap_password').send_keys('Welcome1')
#
# # Keep me signed in
# driver.find_element(By.NAME, 'rememberMe').click()
#
# # Click on sign in button
# driver.find_element(By.ID, 'signInSubmit').click()
# time.sleep(15)
#
# # Enter Otp
#
# # Don't require OTP on this browser
# driver.find_element(By.ID, 'auth-mfa-remember-device').click()
#
# # Click on sign in
# driver.find_element(By.ID, 'auth-signin-button').click()
# time.sleep(10)
#
# # Navigate to the Help page
# driver.get("https://sellercentral.amazon.com/help/center?mons_redirect=stck_reroute")
#
# # Clicking on Inventory lost in FBA warehouse
# driver.find_element(By.CLASS_NAME, 'description meld-label-description').click()
#
# # Enter FNSKU
# driver.find_element(By.TAG_NAME, 'input').click()
#
#
# # You can continue interacting with the page or perform additional actions here
#
# # Wait for user input before exiting
# input("Press Enter to close the browser...")
#
# # Close the browser when the user presses Enter
# driver.quit()
