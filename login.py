from selenium import webdriver
import time
driver = webdriver.Edge()
driver.get("https://itsm.fenetwork.com/HPSM9.33_PROD/index.do")
time.sleep(2)
driver.find_element_by_id("LoginUsername").send_keys("55489")
driver.find_element_by_id("LoginPassword").send_keys("First internship27")
driver.find_element_by_id("loginBtn").click()

