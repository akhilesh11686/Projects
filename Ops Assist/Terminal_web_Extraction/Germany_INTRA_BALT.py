
import requests
import pandas as pd

from dateutil import parser

import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning


import datetime
import pytz


class INTRA_BALT:

    def convert_datetime_timezone(dt,tz1,tz2):
        tz1 = pytz.timezone(tz1)
        tz2 = pytz.timezone(tz2)
        

        dt = datetime.datetime.strptime(dt,"%Y-%m-%d %H:%M:%S")
        dt = tz1.localize(dt)
        dt = dt.astimezone(tz2)
        dt = dt.strftime("%Y-%m-%d %H:%M:%S")
        return dt


    def Germany_BALT():
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
        url = 'https://coast.hhla.de/api/execute-report/Standard-Report-Segelliste'
        res = requests.get(url, verify=False)
        val= res.json()
        # spd = pd.DataFrame.to_json(val)

        dfN = pd.DataFrame()

        appended_data = []

        # for ky, itm in val.items():
        col = val['resultTables']
        for k in range(len(col)):
                col1 = col[0]['rows']
                for i in range(len(col1)):
                    col2 = col1[i]
                    for k in range(len(col2)):
                        outVal=col2[k]['value']
                        test_str = outVal
                        if test_str != None:
                            format = "%Y-%m-%d"
                            try:
                                if parser.parse(test_str):
                                    test_str = test_str.replace('T'," ").replace('+02:00',"").replace('+01:00',"")
                                    outVal =INTRA_BALT.convert_datetime_timezone(test_str, "CET", "Asia/Kolkata")
                                    appended_data.append(outVal)
                            except ValueError:
                                appended_data.append(outVal)
                        else:
                            appended_data.append(outVal)

                    if len(appended_data)!=0:
                        df = pd.DataFrame([appended_data],columns=['departure','departure (scheduled)','arrival','arrival (Scheduled)','export voyage','callsign','import voyage','###','start of loading','end of loading','start of discharge','end of discharge','vessel name','type of vessel','terminal'])
                        dfN = pd.concat([dfN,df])
                        appended_data.clear()

        # spd = pd.DataFrame(dfN)
        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        dfN.to_excel(file_name,index=False,sheet_name='Germany_INTRA_BALT')
        file_name.save()            
        return