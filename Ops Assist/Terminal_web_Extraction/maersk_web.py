import warnings
import enum

import pandas as pd
import time
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from win32ui import CreateListCtrl

from bs4 import BeautifulSoup
import urllib
import webbrowser

import re
from urllib.request import urlopen
from urllib.request import urlretrieve
from urllib.parse import urlencode
import os

class Mrsk_web():    
    
    def getWebData(vslName,vsl,frmDate):
        # global appended_data
        values = {
            'vesselCode':vsl,
            'fromDate': frmDate,
        }

        qstr = urlencode(values)
        url = 'https://www.maersk.com/schedules/vesselSchedules?' + qstr

        options = Options()
        # options.add_argument("--window-size=1920,1200")
        options.add_argument("--headless")  # temp
        os.environ['WDM_SSL_VERIFY']='0'    #Disable the SSL
        driver = webdriver.Chrome(ChromeDriverManager().install(),options = options)

        # url = 'https://www.maersk.com/schedules/vesselSchedules?vesselCode=997&fromDate=2022-02-03'
        # driver.get('https://www.maersk.com/schedules/vesselSchedules')
        driver.get(url)
        time.sleep(5)

        # accptAll = driver.find_element_by_xpath('//*[@id="coiPage-1"]/div[2]/button[3]')
        accptAll = driver.find_element('xpath','//*[@id="coiPage-1"]/div[2]/button[3]')

        driver.execute_script("arguments[0].click();", accptAll)

        time.sleep(5)


        dfN = pd.DataFrame()
        dList = []
        rows = driver.find_elements(by=By.CLASS_NAME, value='ptp-results__transport-plan--item')
        # rows = driver.find_elements(by=By.CLASS_NAME, value='vessel-schedules__results')
        for rw in rows:
            dRw = rw.find_elements(by=By.CLASS_NAME, value='font--small')
            for d1 in dRw:
                rwOut = d1.get_attribute('innerHTML')
                if len(rwOut)!=0:
                    if "<" in rwOut:            
                        strVal = d1.get_attribute('innerHTML')
                        res = re.findall(r'>(.*?)</', strVal)
                        outVal =  res[0]
                        dList.append(outVal)
                    else:
                        outVal=d1.get_attribute('innerHTML')
                        dList.append(outVal)

            try:
                df = pd.DataFrame([dList],columns=["a","b","c","d"])
                # dfN =dfN.append(df)
                dfN =pd.concat([dfN,df])
                dList.clear()
            except:
                df = pd.DataFrame([dList],columns=["a","b"])
                # dfN =dfN.append(df)
                dfN =pd.concat([dfN,df])
                dList.clear()

        # dfN =dfN.append(df)
        dfN =pd.concat([dfN,df])

        driver.quit()
        # file_name = 'OutData.xlsx'
        # dfN.to_excel(file_name,index=False)
        dfN['VesselName']=vslName
        return dfN

#%%




