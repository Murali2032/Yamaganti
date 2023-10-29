import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

try:
    service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)

    # Navigate to the initial URL
    driver.get("https://sellercentral.amazon.com/help/center")
    driver.find_element(By.ID, 'ap_email').send_keys('amazonclaims2@gmail.com')
    driver.find_element(By.ID, 'ap_password').send_keys('Welcome1')

    # Keep me signed in
    driver.find_element(By.NAME, 'rememberMe').click()

    # Click on sign in button
    driver.find_element(By.ID, 'signInSubmit').click()
    time.sleep(25)

    # Enter Otp

    # Don't require OTP on this browser
    driver.find_element(By.ID, 'auth-mfa-remember-device').click()

    # Click on sign in
    driver.find_element(By.ID, 'auth-signin-button').click()
    time.sleep(10)

    # Navigate to the Help page
    driver.get("https://sellercentral.amazon.com/help/center?mons_redirect=stck_reroute")

    try:
        # Clicking on Inventory lost in FBA warehouse
        driver.find_element(By.CLASS_NAME, 'title meld-label-message').click()
    except NoSuchElementException:
        print("Element 'Inventory lost in FBA warehouse' not found.")

    # Enter FNSKU
    driver.find_element(By.CSS_SELECTOR, 'i[aria-hidden="true"]').click()

    # You can continue interacting with the page or perform additional actions here

    # Wait for user input before exiting
    input("Press Enter to close the browser...")

finally:
    # Close the browser when done or when an exception occurs
    driver.quit()
