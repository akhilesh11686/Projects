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


class RWG_terminal():
    
    def Rotterdam_RWGt():
        
        url = "https://rwgservices.rwg.nl/Modality/VesselArrivalTimes"
        htm1 = urlopen(url).read()


        data = BeautifulSoup(htm1,'html.parser')

        outVal_tb = data.find_all("table")

        dfN = pd.DataFrame()

        appended_data = []
        for outVal in outVal_tb:
            head_Clmn = outVal.find_all('tr')
            for th in head_Clmn:
                td =th.find_all('td')
                for k in td:
                    if k.string!= None:
                        outVal1 = k.string
                        appended_data.append(outVal1)
                    else:
                        outVal1 = k.string
                        appended_data.append(outVal1)
                        
                
                if len(appended_data)!=0:
                    try:
                        df = pd.DataFrame([appended_data],columns=['ETA','ETD','Object','Operator','Service','Inbound id','Outbound id','Call reference number','Modality'])
                    except ValueError as error:
                        df = pd.DataFrame([appended_data],columns=['ETA','ETD','Object','Operator','Service','Inbound id','Outbound id','Call reference number','Modality'])
                    dfN = pd.concat([dfN,df])
                    appended_data.clear()

        # print("Done")
        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='Rotterdam_RWG')
        file_name.save()            
        