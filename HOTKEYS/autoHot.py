#%%
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import re
import os
import pandas as pd
import time
from bs4 import BeautifulSoup
from tkinter import filedialog
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import random

# from webdriver_manager.chrome import ChromeDriverManager

# driver = Chrome(ChromeDriverManager().install())
# driver = webdriver.Chrome(ChromeDriverManager().install(),options = Options)

def Browse(x):
    pth = filedialog.askopenfilename(title=x)
    return pth
vslList = Browse('Get Vessel Details')
vDf = pd.read_excel(vslList)
dfN1 = pd.DataFrame()

#%%
cosco_list = vDf.loc[vDf['Operator']=='COSCO','Vessel Name'].to_list()
if bool(cosco_list)==True:

    for nVessl in cosco_list: 

        url = 'https://elines.coscoshipping.com/ebusiness/'

        options = Options()
        # options.headless = True
        options.add_argument("--window-size=1920,1200")
        driver = webdriver.Chrome(executable_path=r"./chromedriver", options=options)



        # options = Options()
        # options.add_argument("--window-size=1920,1200")
        # options.add_argument("--headless")  # temp

        os.environ['WDM_SSL_VERIFY']='0'    #Disable the SSL
        # driver = webdriver.Chrome(ChromeDriverManager().install(),options = options)
        driver = webdriver.Chrome(options=options)

        driver.get(url)

        driver.implicitly_wait(20)

        # time.sleep(5)
        # Wait
        
        # WebDriverWait(driver, 100).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

        driver.implicitly_wait(20)
        driver.find_element(By.XPATH,'/html/body/div[3]/div[2]/div/div/div[3]/div/button').click()

        driver.implicitly_wait(20)
        driver.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[1]/ul/li[2]').click()

        driver.implicitly_wait(20)
        driver.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[2]/div/div/div/div/div/div[1]/div/div/div/div/div[3]').click()

        # Clear text field
        driver.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[2]/div/div/div/div/div/div[2]/div[2]/div/form/div/div[1]/div/div/div/div[1]/input').clear()
        # vslName = 'COSCO SHIPPING ANDES'
        vslName = nVessl
        driver.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[2]/div/div/div/div/div/div[2]/div[2]/div/form/div/div[1]/div/div/div/div[1]/input').send_keys(vslName)
        # driver.find_elements(By.CLASS_NAME,'ivu-select-dropdown-list')
        # time.sleep(1)

        # Wait
        driver.implicitly_wait(20)
        # WebDriverWait(driver, 90000).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')


        kk = driver.find_elements(By.CLASS_NAME,'ivu-select-dropdown-list')
        for i in kk:
            try:        
                i.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[2]/div/div/div/div/div/div[2]/div[2]/div/form/div/div[1]/div/div/div/div[2]/ul[2]/div/li').click()        
                time.sleep(1)
                driver.find_element(By.XPATH,'//*[@id="header"]/div/div[2]/div[3]/div[2]/div/div/div/div/div/div[2]/div[2]/div/form/div/div[2]/div/button').click()
                break
            except Exception:
                pass

        # Wait
        driver.implicitly_wait(20)

        # VslName
        vslPage = driver.find_element(By.XPATH,'/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div[1]/div/div/div/div[1]')
        vSoup = BeautifulSoup(vslPage.get_attribute('innerHTML'),'html.parser')
        vessleName = vSoup.find('span').getText()


        # Get table content
        clmHead1 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[1]')
        clmHead1_text = BeautifulSoup(clmHead1.get_attribute('innerHTML'),'html.parser').getText()

        clmHead2 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[2]')
        clmHead2_text = BeautifulSoup(clmHead2.get_attribute('innerHTML'),'html.parser').getText()

        clmHead3 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[3]')
        clmHead3_text = BeautifulSoup(clmHead3.get_attribute('innerHTML'),'html.parser').getText()

        clmHead4 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[4]')
        clmHead4_text = BeautifulSoup(clmHead4.get_attribute('innerHTML'),'html.parser').getText()

        clmHead5 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[5]')
        clmHead5_text = BeautifulSoup(clmHead5.get_attribute('innerHTML'),'html.parser').getText()

        clmHead6 = driver.find_element(By.XPATH,'//*[@id="downloadSaislingSchedule"]/div[6]')
        clmHead6_text = BeautifulSoup(clmHead6.get_attribute('innerHTML'),'html.parser').getText()

        # Get columns
        df = pd.DataFrame(columns=[clmHead1_text,clmHead2_text,clmHead3_text,clmHead4_text,clmHead5_text,'Vesel_Name','Voyage'])

        # Wait
        driver.implicitly_wait(20)
        # WebDriverWait(driver, 100).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

        pgCnt = driver.find_element(By.XPATH,'//*[@id="capture"]/div[2]/div[2]')
        soup = BeautifulSoup(pgCnt.get_attribute('innerHTML'), 'html.parser')


        # Wait
        driver.implicitly_wait(20)
        # WebDriverWait(driver, 100).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')

        tList = []
        for i in soup:
            try:
                vyg = i.find('span').text
                tbl = i.findAll('tr')
                for j in tbl:
                    cData = j.findAll('td')
                    for k in cData:
                        outV = k.text
                        if outV.find('No filter data') !=-1:
                            break            
                        tList.append(k.text)         
                    tList.append(vessleName)
                    tList.append(vyg)
                    df= df.append( pd.DataFrame([tList],columns=[clmHead1_text,clmHead2_text,clmHead3_text,clmHead4_text,clmHead5_text,clmHead6_text,'Vesel_Name','Voyage']))
                    tList.clear()
            except Exception:
                tList.clear()
        driver.back()
        driver.implicitly_wait(20)
        driver.close()
        dfN1 = pd.concat([df,dfN1])
dfN1.to_excel('2.xlsx',index=False) 
#end cosco
#                
#%%
maersk_list = vDf.loc[vDf['Operator']=='MAERSK','Vessel Name'].to_list()
if bool(maersk_list)==True:
    for nVessl in maersk_list: 
        url = 'https://www.maersk.com/schedules/vesselSchedules'

        options = Options()
        # options.headless = True
        options.add_argument("--window-size=1920,1200")
        driver = webdriver.Chrome(executable_path=r"./chromedriver", options=options)

        driver.get(url)

        driver.implicitly_wait(10)  
        # Accept the cookies
        driver.find_element(By.CSS_SELECTOR,'#coiPage-1 > div.coi-banner__page-footer > button.coi-banner__accept.coi-banner__accept--fixed-margin').click()
        driver.implicitly_wait(20)    
        # moves = [Keys.LEFT, Keys.DOWN, Keys.RIGHT, Keys.UP,Keys.ENTER]
        moves = [Keys.DOWN]
        driver.find_element(By.XPATH,'//*[@id="vesselName"]').send_keys(nVessl)

        # body.send_keys(Keys.ARROW_DOWN)
        # body.send_keys(Keys.ENTER) 
        driver.implicitly_wait(20)




        # driver.implicitly_wait(20)


print('Done!!!')
#%%
