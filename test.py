# # from selenium import webdriver
# # from selenium.webdriver.chrome.service import Service
# #
# # service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
# # options = webdriver.ChromeOptions()
# # driver = webdriver.Chrome(service=service, options=options)
# #
# # # Navigate to a URL
# # driver.get("https://www.youtube.com/watch?v=_3latQQYa74&ab_channel=TV9TeluguLive")
# #
# # # Assuming the video has an 'play' button that you want to click
# # like_button = driver.find_element_by_xpath("//*[@id="segmented-like-button"]/ytd-toggle-button-renderer/yt-button-shape/button/yt-touch-feedback-shape/div/div[2]").click
# #
# # # Wait for a while (optional)
# # time.sleep(10)
# #
# # # Quit the WebDriver when done
# # driver.quit()
#
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# import time
#
#
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# import time
#
# service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
# options = webdriver.ChromeOptions()
# driver = webdriver.Chrome(service=service, options=options)
#
# # Navigate to a URL
# driver.get("https://www.youtube.com/watch?v=_3latQQYa74&ab_channel=TV9TeluguLive")
# driver.maximize_window()
# time.sleep(4)
#
# # Find the 'like' button element by its XPath
# like_button = driver.find_element_by_xpath("//div[@class='yt-spec-touch-feedback-shape__fill']")
#
# # Click the 'like' button
# like_button.click()
#
# # Wait for a while (optional)
# time.sleep(5)  # Adjust the sleep time as needed
#
# # Get the state of the 'like' button after clicking
# new_state = like_button.get_attribute("aria-pressed")
#
# # Check if the state changed after clicking
# if initial_state != new_state:
#     print("Successfully clicked the 'like' button.")
# else:
#     print("Failed to click the 'like' button.")
#
# # Quit the WebDriver when done
# driver.q

#===============================

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time

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

# This Element is inside a single shadow DOM.
cssSelectorForHost1 = ".icon.icon-selected"
time.sleep(1)

# Find the host element
hostElement = driver.find_element(By.CSS_SELECTOR, cssSelectorForHost1)

# Get the shadow DOM
shadow = driver.execute_script("return arguments[0].shadowRoot", hostElement)

time.sleep(1)

# Find the element inside the shadow DOM
elementInsideShadow = shadow.find_element(By.CSS_SELECTOR, "i[aria-hidden='true']")

# Now you can interact with elementInsideShadow

# Click the host element
hostElement.click()

# Rest of your code

input("Press Enter to close the browser...")

