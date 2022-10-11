#%%
import pandas as pd
import os
import selenium
from selenium import webdriver
import time
from PIL import Image
import io
import requests
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from webdriver_manager.microsoft import IEDriverManager



options = Options()
options.add_argument("--window-size=1300,1000")
os.environ['WDM_SSL_VERIFY']='0'    #Disable the SSL
#Install Driver

# https://simpleit.rocks/python/selenium-webdriver-exception-executable-in-path-error/

try:
    driver = webdriver.Chrome(ChromeDriverManager().install())
except ConnectionError:
    driver = webdriver.Ie(IEDriverManager().install())

#Specify Search URL 
search_url="https://www.msc.com/en/search-a-schedule" 

driver.get(search_url)

time.sleep(10)

vslLink = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/ul/li[2]/button')
vslLink.click()

vslList = ['MSC SINDY','MSC TOKYO','MSC TORONTO','MSC LUCY']
dfN = pd.DataFrame()
xlList = []

for vsl  in vslList:
    vslNme = driver.find_element_by_xpath('//*[@id="vessel"]')
    vslNme.clear()
    vslNme.send_keys(vsl)
    vslNme.send_keys(Keys.TAB)
    time.sleep(5)
    vslNme.send_keys(Keys.ENTER)

    
    rslt = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/form[2]/div[2]/button')
    rslt.click()

    time.sleep(10)

    tblData = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/div[1]/div/div')
    # Parsing Html
    soup = BeautifulSoup(tblData.get_attribute('innerHTML'), 'lxml')
    spans = soup.findAll('span')

    [xlList.append(span.text) for span in spans]
    df = pd.DataFrame(xlList)
    df['VesselName']=vsl
    dfN = pd.concat([dfN,df])


file_name = pd.ExcelWriter('Partener_Vessels.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
dfN.to_excel(file_name,index=False,sheet_name='Msc')
file_name.save()                   
# dfN.to_excel('dd.xlsx')
print('Done')
driver.quit()
