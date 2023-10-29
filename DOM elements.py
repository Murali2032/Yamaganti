import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Navigate to the initial URL
driver.get("https://sellercentral.amazon.com/help/center")
time.sleep(40)
shopnow=driver.execute_script('document.querySelector("#root > div > section.kat-col.body > section > kat-box > div.meld-transcript > div:nth-child(1) > div > div > div > div > div > div > div.card-link-button > kat-card > div > div.title-row > div.icon-container > kat-icon").shadowRoot.querySelector("i")')
driver.execute_script('arguments[0].click();', shopnow)

input("Press Enter to close the browser...")

# driver.get("https://shop.polymer-project.org/")
# shopnow=driver.execute_script('return document.querySelector("body > shop-app").shadowRoot.querySelector("iron-pages > shop-home").shadowRoot.querySelector("div:nth-child(2) > shop-button > a")')
# # driver.execute_script('arguments[0].click();', shopnow)


