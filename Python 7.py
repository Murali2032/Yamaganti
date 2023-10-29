# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.options import Options
# from datetime import datetime
#
# service = Service(executable_path='C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe')
# options = webdriver.ChromeOptions()
# driver = webdriver.Chrome(service=service, options=options)
#
# chrome_options = Options()
# chrome_options.add_experimental_option("debuggerAddress","localhost:7777")
#
# driver.get("https://www.google.com/")



from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# chrome_options = Options()
# chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# #Change chrome driver path accordingly
# chrome_driver = "(C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe)"
# driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
# print (driver.title)
#
# driver = webdriver.Chrome(executable_path='C:\Program Files\Chrome Driver\chromedriver.exe')




#
# import os
#
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.options import Options
# from datetime import datetime
#
# chrome_options = Options()
# chrome_options.add_experimental_option("debuggerAddress", "localhost:8989")
# # chrome_options.add_argument("--headless")
# chrome_options.add_argument("--incognito")
#
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)
# driver.maximize_window()
#
# driver.get("https://www.google.com/")
# now = datetime.now()
# current_time = now.strftime("%H_%M_%S")
# driver.get_screenshot_as_file(os.getcwd()+'/screenshots/Screenshot_'+current_time+'.png')
#
# driver.get("https://www.facebook.com/")
# now = datetime.now()
# current_time = now.strftime("%H_%M_%S")
# driver.get_screenshot_as_file(os.getcwd()+'/screenshots/Screenshot_'+current_time+'.png')
#
# # driver.quit()



System.setProperty("webdriver.chrome.driver","C:/Users/YamagantiMuraliKrish/PycharmProjects/Selenium_01/chromedriver.exe");
ChromeDriver driver=new ChromeDriver();
Capabilities cap=driver.getCapabilities();
Map<String, Object> myCap=cap.asMap();
System.out.println(myCap);
