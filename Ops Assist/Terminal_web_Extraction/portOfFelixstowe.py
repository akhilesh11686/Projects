#%%
from cgitb import html
import urllib.request
from uuid import RESERVED_FUTURE
from bs4 import BeautifulSoup
import pandas as pd 


class uk_Felixstowe():

    def terminal_Felixstowe():
        with urllib.request.urlopen('https://www.portoffelixstowe.co.uk/sailing-schedule/shipping-information/') as response:
            html = response.read().decode('utf-8')
            soup = BeautifulSoup(html)

        outVal_tb = soup.find_all("table")

        dfN = pd.DataFrame()
        appended_data = []
        for outVal in outVal_tb:
            head_Clmn = outVal.find_all('tr')
            for trs in head_Clmn:

                td =trs.find_all('td')
                if len(td) == 0:
                    ths_ =trs.find_all('th')
                    for x in ths_:
                        appended_data.append(x.getText())                
                else:
                    for k in td:
                        appended_data.append(k.getText())
                        
                if len(appended_data)!=0:
                    try:
                        df = pd.DataFrame([appended_data],columns=['A','B','C','D','E','F','G','H','I'])
                    except ValueError as error:
                        df = pd.DataFrame([appended_data],columns=['A','B','C','D','E','F','G'])

                    dfN = pd.concat([dfN,df])
                    appended_data.clear()
                    
                        
        # print("Done")
        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='PortOfFelixstowe')
        file_name.save()            
        return
   



