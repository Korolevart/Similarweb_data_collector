import os
os.chdir('J:/MEC/Analytics and Insight/Korolev/Python/Similarweb')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from openpyxl import load_workbook
from time import sleep
import urllib
import urllib.request
import json
import pandas as pd
from pandas.io.json import json_normalize

start_date = '2015|12|01'
end_date = '2015|12|31'

login_url = 'https://www.similarweb.com/account/login'
request_url ='https://pro.similarweb.com/api/websiteanalysis/GetTrafficSourcesOverviewNew?country=999&from='
text1 = '&isWWW=false&isWindow=false&key='
text2 = '&to='

driver = webdriver.Firefox()

driver.get(login_url)

username = driver.find_element_by_name("email")
username.send_keys("********")

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

data_frame = []

site = url_values[2]
for site in url_values:
    full_url = request_url + start_date + text1 + site + text2 + end_date
    driver.get(full_url)
    element=driver.find_element_by_xpath(("//html/body/pre"))
    response=element.get_attribute('innerHTML')
    jsonResponse=json.loads(response)
    jsonData = jsonResponse[site]
    jsonData = jsonData['Volumes']
    Appstore_Internals_Organic = jsonData['Appstore Internals']['Organic'][0]
    Appstore_Internals_Paid = jsonData['Appstore Internals']['Paid'][0]
    Direct_Organic = jsonData['Direct']['Organic'][0]
    Direct_Paid = jsonData['Direct']['Paid'][0]
    Mail_Organic = jsonData['Mail']['Organic'][0]
    Mail_Paid = jsonData['Mail']['Paid'][0]
    Paid_Referrals_Organic = jsonData['Paid Referrals']['Organic'][0]
    Paid_Referrals_Paid = jsonData['Paid Referrals']['Paid'][0]
    Referrals_Organic = jsonData['Referrals']['Organic'][0]
    Referrals_Paid = jsonData['Referrals']['Paid'][0]
    Search_Organic = jsonData['Search']['Organic'][0]
    Search_Paid = jsonData['Search']['Paid'][0]
    Social_Organic = jsonData['Social']['Organic'][0]
    Social_Paid = jsonData['Social']['Paid'][0]
    Site = site
    Lists=[]
    List = [Site, Appstore_Internals_Organic, Appstore_Internals_Paid, Direct_Organic, Direct_Paid, Mail_Organic, Mail_Paid, Paid_Referrals_Organic,Paid_Referrals_Paid,Referrals_Organic,Referrals_Paid,Search_Organic,Search_Paid,Social_Organic,Social_Paid]
    #data = json_normalize(jsonData)
    #data['Site']=site
    data_frame.append(List)

data = pd.DataFrame(data_frame)
data.columns = ['Site', 'Appstore_Internals_Organic', 'Appstore_Internals_Paid', 'Direct_Organic', 'Direct_Paid', 'Mail_Organic', 'Mail_Paid', 'Paid_Referrals_Organic','Paid_Referrals_Paid','Referrals_Organic','Referrals_Paid','Search_Organic','Search_Paid','Social_Organic','Social_Paid']

writer = pd.ExcelWriter('J:/MEC/Analytics and Insight/Korolev/Python/Similarweb/report.xlsx')
data.to_excel(writer,'Sheet1')
writer.save()

