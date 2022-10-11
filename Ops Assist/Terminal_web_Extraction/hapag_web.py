
# #%%
# from operator import index
# from pickle import FALSE
# from socket import SOL_UDP
# import warnings
# import enum
# from xml.etree import ElementTree
# from docx.table import _Row
# import pandas as pd
# import time
# from bs4 import BeautifulSoup
# from selenium.webdriver.chrome.options import Options
# from selenium import webdriver
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.common.keys import Keys
# from selenium.common.exceptions import TimeoutException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.by import By
# from win32ui import CreateListCtrl
# from selenium.webdriver.support.ui import Select
# import urllib
# import webbrowser

# import re
# from urllib.request import urlopen
# from urllib.request import urlretrieve
# from urllib.parse import urlencode
# import os

# class hapag_l_web():

#     def getHapag(vslName):

#         url = 'https://www.hapag-lloyd.com/en/online-business/track/vessel-tracker-solution.html'

#         options = Options()
#         options.add_argument("--window-size=1300,1000")
#         # options.add_argument("--window-size=1920,1200")
#         # options.add_argument("--headless")
#         os.environ['WDM_SSL_VERIFY']='0'    #Disable the SSL
#         driver = webdriver.Chrome(ChromeDriverManager().install(),options = options)

#         driver.get(url)
#         time.sleep(10)


#         accptAll = driver.find_element_by_xpath('//*[@id="accept-recommended-btn-handler"]')
#         driver.execute_script("arguments[0].click();", accptAll)

#         time.sleep(10)

#         # select drop down list
#         # https://intellipaat.com/community/4266/how-to-select-a-drop-down-menu-option-value-with-selenium-python
#         # select = Select(driver.find_element_by_xpath('//*[@id="ext-gen129"]'))

#         # # select by visible text
#         # select.select_by_visible_text('LIVORNO EXPRESS')

#         vslNme = driver.find_element_by_xpath('//*[@id="ext-gen129"]')
#         vslNme.send_keys(vslName)
#         vslNme.send_keys(Keys.ENTER)

#         time.sleep(10)
#         sbmtBtn = driver.find_element_by_xpath('//*[@id="schedules_vessel_tracing_f:hl24"]')
#         driver.execute_script("arguments[0].click();", sbmtBtn)



#         dfN= pd.DataFrame()
#         th_list = []

#         # =========================1st table start

#         tblData = driver.find_element_by_xpath('//*[@id="schedules_vessel_tracing_f:hl40"]')
#         # Parsing Html
#         soup = BeautifulSoup(tblData.get_attribute('innerHTML'), 'lxml')

#         thlist = soup.find('thead').find('tr').find_all('th')
#         for thT in thlist:
#             th_list.append(thT.text)
#         df = pd.DataFrame([th_list])
#         dfN = pd.concat([df,dfN])
#         th_list.clear()

#         thlist = soup.find('tbody').find_all('tr')
#         for trV in thlist:
#             tdCnt = trV.find_all('td')
#             for tdT in tdCnt:
#                 th_list.append(tdT.text)
#             df = pd.DataFrame([th_list])
#             th_list.clear()
#             dfN = pd.concat([dfN,df])

#         dfN.to_excel("OutData.xlsx")
#         # ========================= 1st table end
#         dfN = pd.DataFrame(None)


#         tblData1 = driver.find_element_by_xpath('//*[@id="schedules_vessel_tracing_f:hl68"]')
#         soup1 = BeautifulSoup(tblData1.get_attribute('innerHTML'), 'lxml')
#         thlist1 = soup1.find('thead').find('tr').find_all('th')
#         for thT in thlist1:
#             th_list.append(thT.text)
#         df = pd.DataFrame([th_list])
#         dfN = pd.concat([df,dfN])
#         th_list.clear()


#         thlist1 = soup1.find('tbody').find_all('tr')
#         for trV in thlist1:
#             tdCnt1 = trV.find_all('td')
#             for tdT1 in tdCnt1:
#                 th_list.append(tdT1.text)
#                 df = pd.DataFrame([th_list])
#             th_list.clear()
#             dfN = pd.concat([dfN,df])

#         df2 = pd.read_excel('OutData.xlsx')
#         dfN = pd.concat([dfN,df2])

#         driver.quit()

#         # dfN.to_excel("OutData.xlsx")
#         dfN['VesselName']=vslName
#         return dfN


#%%

import xlwings as xw
from openpyxl import load_workbook
import pandas as pd


class hapag_l_web():

    def getHapag():

        VslN = pd.read_excel('Operator wise vessel details.xlsx',sheet_name="vessel & service data")
        vlName = VslN[VslN['Operator'] == 'Hapag']

        unqVsl = vlName['Vessel Name'].drop_duplicates()
        strList = unqVsl.to_list()

        strVal = "|".join(str(k) for k in strList)

        wb = xw.Book('Hpag.xlsm')
        mRun = wb.macro('Automate_IE_Load_Page')
        mRun(strVal)
        wb.save()
        wb.close()


        path = './Hpag.xlsm'
        sheet_name = 'Sheet1'
        Rslt = load_workbook(path,keep_vba=True)
        df = pd.DataFrame(Rslt['Sheet1'].values)

        file_name = pd.ExcelWriter('Partener_Vessels.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        df.to_excel(file_name,index=False,sheet_name='Hapag')
        file_name.save()




#%%

url = 'https://www.hapag-lloyd.com/en/online-business/track/vessel-tracker-solution.html'
