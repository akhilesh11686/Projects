#%%
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
from win32ui import CreateListCtrl


class dpLondon():

    def london_Gateway():
        options = Options()
        # options.add_argument("--window-size=1920,1200")
        options.add_argument("--headless")
        
        # driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
        # driver = webdriver.Chrome(ChromeDriverManager().dont_verify_ssl().install())
        driver = webdriver.Chrome(ChromeDriverManager().install(),options = options)

        driver.get('https://www.dpworld.com/london-gateway/port/vessel-schedule')
        time.sleep(10)

        try:
            element_present = EC.visibility_of_element_located((By.XPATH, '//*[@id="schedule-table"]/table'))
            main = WebDriverWait(driver, 10).until(element_present)

            wait = WebDriverWait(driver, 10)
            ele = driver.find_element_by_xpath('//*[@id="table-loadmore"]/a')

            # site for click issue
            # https://stackoverflow.com/questions/57741875/selenium-common-exceptions-elementclickinterceptedexception-message-element-cl
            driver.execute_script("arguments[0].click();", ele)

            # wait = WebDriverWait(driver, 40)
            time.sleep(30) 

            rows = driver.find_elements_by_xpath('//*[@id="schedule-table"]/table/tbody/tr')    
            # print(len(rows))


            dfN = pd.DataFrame()
            appended_data = []

            for rw in rows:
                splVal = rw.get_attribute('innerHTML').replace('\n','').replace('<td>','').replace('</td>','|')
                lst = splVal.split('|')
                                
                if len(lst)!=0:
                    lst = [x.strip(' ') for x in lst]
                    df = pd.DataFrame([lst],columns=['Vessel Name','Operator','Service','Phase','ETA','ETD','of call','call',''])
                    dfN =dfN.append(df)

        except TimeoutException:
            driver.quit()

        # print("Done")

        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='DpWorld_London')
        file_name.save()            
        return

# %%
