
#%%
import requests
from ast import literal_eval
import pandas as pd
import json
import openpyxl
import urllib
import webbrowser
from bs4 import BeautifulSoup, element
import io
# urllib2
from urllib.request import urlopen
from urllib.request import urlretrieve
from urllib.parse import urlencode

import http


class T_913():
    def ant_t913():
        # url = "https://info.eworld1700.be/vessellist"
        url = "https://www.psa-antwerp.be/nl/noordzee-terminal-q913-yard-opening-times"
        htm1 = urlopen(url).read()


        data = BeautifulSoup(htm1,'html.parser')

        # print(data.prettify())
        outVal_tb = data.find_all("table")

        dfN = pd.DataFrame()
        # head_Clmn = outVal.find_all('th')
        # for th in head_Clmn:
        #     print(th.string)
        appended_data = []
        for outVal in outVal_tb:
            head_Clmn = outVal.find_all('tr')
            for th in head_Clmn:
                td =th.find_all('td')
                for k in td:
                    if k.string!= None:
                        outVal1= ''.join(k.string.split())
                        appended_data.append(outVal1)
                    else:
                        outVal1 =k.find('p').string
                        appended_data.append(outVal1)
                        
                
                if len(appended_data)!=0:
                    try:
                        df = pd.DataFrame([appended_data],columns=['VESSEL CODE','VESSEL NAME','VOYAGE OUT','ETA','YARD OPENING TIME','CHANGED'])
                    except ValueError as error:
                        df = pd.DataFrame([appended_data],columns=['VESSEL CODE','VESSEL NAME','VOYAGE OUT','ETA','YARD OPENING TIME'])
                    # result_dataframe = dfN.append(df)
                    dfN = pd.concat([dfN,df])
                    appended_data.clear()

        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='Antwerp_913')
        file_name.save()    
        return