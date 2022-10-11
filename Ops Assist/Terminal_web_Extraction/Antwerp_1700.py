# pip install webdriver-manager  ->>to get lates version of selenium

# from sys import last_value
import warnings
import enum
import pandas as pd
import time

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class T_1700():
    def ant_t1700():

        options = Options()
        options.add_argument("--window-size=1920,1200")
        # options.add_argument("--headless")

        driver = webdriver.Chrome(ChromeDriverManager().install(),options = options)

        driver.get('https://info.eworld1700.be/vessellist')
        

        time.sleep(10)
        # skeep Deprecated warnig
        warnings.filterwarnings("ignore", category=DeprecationWarning)

        frstPart= ""
        sndpart =""
        thrPart=""
        try:
            element_present = EC.presence_of_element_located((By.ID, "vessellist"))
            main = WebDriverWait(driver, 10).until(element_present)
            # artl = driver.find_elements_by_tag_name('tr')
            # eNum = len(driver.find_elements_by_tag_name('tr'))+1

            # rwCnt = len(driver.find_element_by_xpath('//*[@id="vessellist"]/table/tbody/tr').text)
            rwCnt = len(driver.find_elements_by_tag_name('tr'))
            clmCnt = len(driver.find_element_by_xpath('//*[@id="vessellist"]/table/tbody/tr[1]/th').text)

            # frstPart = "'//*[@id='vessellist']/table/tbody/tr["
            # sndpart = "]/th["
            # thrPart = "]')"

            # strVal = driver.find_element_by_xpath(
            # driver.find_element_by_xpath('//*[@id="vessellist"]/table/tbody/tr[1]/th[1]').text
            # 'Vessel Name'
            # driver.find_element_by_xpath('//*[@id="vessellist"]/table/tbody/tr[2]/td[1]').text
            # 'X-PRESS ANNAPURNA'

            dfN = pd.DataFrame()

            appended_data = []
            # df = pd.DataFrame([appended_data],columns=['Vessel Name','Vessel Code','Voy IN','Voy OUT','Cargo Opening','ETA','ATA','ETD','Service','Vessel Status'])
            for n in range(2,rwCnt):
                # print(appended_data)
                # df.loc[len(df)] = [appended_data]
                
                for m in range(1,clmCnt):
                    if n == 1:
                        frstPart = "//*[@id='vessellist']/table/tbody/tr["
                        sndpart = "]/th["
                        thrPart = "]"         

                        fnlPath = frstPart +str(n)+ sndpart +str(m)+ thrPart
                        outVal = driver.find_element_by_xpath(fnlPath).text
                        appended_data.append(outVal)

                    elif n>1:
                        frstPart = "//*[@id='vessellist']/table/tbody/tr["
                        frstChck = "]"

                        sndpart = "]/td["
                        thrPart = "]"         
                        firstValidation = frstPart +str(n)+frstChck
                        outNum = driver.find_element_by_xpath(firstValidation).text
                        if len(outNum)!=0:
                            fnlPath = frstPart +str(n)+ sndpart +str(m)+ thrPart
                            outVal = driver.find_element_by_xpath(fnlPath).text
                            appended_data.append(outVal)
                        else:
                            break           

                if len(appended_data)!=0:
                    df = pd.DataFrame([appended_data],columns=['Vessel Name','Vessel Code','Voy IN','Voy OUT','Cargo Opening','ETA','ATA','ETD','Service','Vessel Status'])
                    # result_dataframe = dfN.append(df)
                    dfN = pd.concat([dfN,df])
                    appended_data.clear()
            # appended_data = []
            # for i in range(eNum-1):
            #     out = driver.find_elements_by_tag_name('tr')[i].text
            #     if len(out)!= 0:
            #         lsVal = out.split(' ')
            #         appended_data.append(lsVal)
        except TimeoutException:
            pass

        # print(driver.title)
        driver.quit()
        # print("Done")
        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='Antwerp_1700')
        file_name.save()    
        return