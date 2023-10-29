import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Navigate to a URL
driver.get("https://sellercentral.amazon.com/help/center")
driver.find_element(By.ID, 'ap_email').send_keys('amazonclaims@elainandwayne.com')
driver.find_element(By.ID, 'ap_password').send_keys('Welcome1')
driver.find_element(By.ID, 'signInSubmit').click()
time.sleep(25)

# Navigating to the Help page
driver.find_element(By.XPATH, ('//p[@class='description meld-label-description']').click()

# Clicking on Inventory lost in FBA
driver.find_element(By.CLASS_NAME, 'description meld-label-description').click()


# Wait for user input before exiting
input("Press Enter to close the browser...")

# Close the browser when the user presses Enter
driver.quit()
