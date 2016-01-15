import os
os.chdir('J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/NEW')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
from openpyxl import load_workbook
from time import sleep
import urllib
import urllib.request
import json
import pandas as pd
from pandas.io.json import json_normalize


login_url = 'https://www.similarweb.com/account/login'
request_url ='https://pro.similarweb.com/website/analysis/#/'
text1 = '/*/999/2015.12-2015.12/audience/overview?selectTrendLine=visits&aggDuration=monthly'

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.dir", os.getcwd())
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

driver = webdriver.Firefox(profile)

driver.get(login_url)

username = driver.find_element_by_name("email")
username.send_keys("***********")

password = driver.find_element_by_name("password")
password.send_keys("********")

button = driver.find_element_by_xpath(("//button[contains(text(),'Sign In')]"))
button.submit()


wb = load_workbook(filename='J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/Similarweb.xlsx', read_only=True)
ws = wb['Sheet1']
url_values=[]
for row in ws.rows:
    for cell in row:
        url_values.append(cell.value)

sleep(5)


for site in url_values:
    full_url = request_url + site + text1
    print(full_url)
    driver.get(full_url)
    sleep(10)
    driver.find_element_by_xpath("//div[@class='swButton swButton--white export-btn sw-icon-download-new']").click()
    sleep(10)
    driver.find_element_by_xpath("//a[contains(text(),'Download Excel')]").click()
    print("DONE")
    sleep(3)
    os.chdir('J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/NEW')
    #FIND MOST RECENT FILE IN (YOUR) DIR AND RENAME IT
    files = filter(os.path.isfile, os.listdir("J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/NEW"))
    print(files)
    files = [os.path.join("J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/NEW", f) for f in files]
    print(files)
    files.sort(key=lambda x: os.path.getmtime(x))
    print(os.path.getmtime('J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/NEW/Monthly'))
    newest_file = files[-1]
    print(files[-1])
    print(newest_file)
    os.rename(newest_file, site + ".xlsx")
