
#%%
from email import header
from importlib.resources import contents
from time import sleep
import urllib.request
import pandas as pd
import json
import openpyxl
import time

# urllib2
from urllib.request import urlopen,Request
from urllib.request import urlretrieve
from urllib.parse import urlencode
import urllib,json
import urllib.parse

from pydantic import UrlHostTldError
from urllib.request import urlopen, HTTPError, URLError

class Cosco_web_e:    
#     # vsl,frmDate,toDate = 997,'24-12-2021','04-02-2022'
    
    def cosco_getWebData(vslName,vsl,frmDate):
        # vslName = 'COSCO SHIPPING LEO'
        # vsl = 'CNF'
        # frmDate = '28'
        # global appended_data
        values = {
            'vesselCode':vsl,
            'period': frmDate
            # 'toDate': toDate
        }
        qstr = urlencode(values)
        url = 'https://elines.coscoshipping.com/ebschedule/public/purpoShipment/vesselCode?' + qstr

        try:
            # req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            req = Request(url)
            html1 = urlopen(req)            
            html = html1.read() 
            # time.sleep(5)       
            data = json.loads(html)            # Get value of list            
        except HTTPError as e:
            print(qstr + ' HTTP Error code: ', e.code)
        except URLError as e:
            print('URL Error: ', e.reason)
        else:
            # Get value of list
            # print("No Err")
            srcData = data['data']['content']['data']
            spd = pd.DataFrame.from_dict(srcData)
            spd['VesselName']=vslName
            return spd

# #%%
# #########################CMA data Extraction########################################
# # https://www.cma-cgm.com/ebusiness/schedules/voyage/detail?VoyageReference=&VesselReference=COSCO+SHIPPING+GEMINI+%3B+COGMI
# from lxml import html
# import requests
# from bs4 import BeautifulSoup

# url = 'https://www.cma-cgm.com/ebusiness/schedules/voyage/detail?VoyageReference=&VesselReference=COSCO+SHIPPING+GEMINI+%3B+COGMI'    

# page = requests.get(url)
# content = html.fromstring(page.content)

# #%%
# import pandas as pd
# import os
# from selenium import webdriver
# import time

# import requests
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from bs4 import BeautifulSoup
# from webdriver_manager.microsoft import IEDriverManager


# options = Options()
# options.add_argument("--window-size=1300,1000")
# # options.add_argument('headless')   #Disable the SSL

# try:
#     driver = webdriver.Chrome(ChromeDriverManager().install())
# except ConnectionError:
#     driver = webdriver.Ie(IEDriverManager().install())


# __author__ = "Ach"
# url = 'https://www.cma-cgm.com/ebusiness/schedules/voyage/detail?VoyageReference=&VesselReference=COSCO+SHIPPING+GEMINI+%3B+COGMI'

# # driver = webdriver.Chrome(executable_path="./driver/chromedriver")
# driver.get(url)
# # content_element = driver.find_elements(By.CLASS_NAME, "last-updated-call")

# # [print(i.text)for i in content_element]
# # content_element = driver.find_elements(By.CLASS_NAME, "future-call")
# # [print(i.text)for i in content_element]

# content_element = driver.find_elements(By.XPATH, '//div[@id="grid"]/div[2]')
# # [print(i.text)for i in content_element]

# content_html = content_element.get_attribute('innerHTML')
# # driver.close()
# #%%


