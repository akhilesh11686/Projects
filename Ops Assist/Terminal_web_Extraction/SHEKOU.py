# #%%
# import requests
# from urllib.parse import urlencode

# import urllib
# import json
# import pandas as pd

# import datetime
# import tkinter.messagebox
# from gInput import getIn

# class shekou_terminals():

#     def getUrl(frmDate,toDate,pg):

#         Payload = {
#             'System': '',
#         'FullName':'',
#         'StartDate': frmDate,
#         'EndDate': toDate,
#         'Service': '',
#         'Ship': '',
#         'vesselName': '',
#         'PageIndex': pg,
#         'PageSize': 30,
#         'SortBy': '',
#         'IsDescending': 'false',
#         }

#         # headers = {'Accept-Language': 'en,en-US;q=0.9','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537.36'}

#         url = 'https://eportapisct.scctcn.com/api/VesselSchedule?'

#         qstr = urlencode(Payload)
#         mUrl = url + qstr
#         return mUrl

#     def getShekou_e():
#         Ndf= pd.DataFrame()
#         pageN = 1

        


#         # date_string = '2018-12-25'
#         format = "%Y-%m-%d"
#         msg = tkinter.messagebox
#         try:
#             # num,name1 = [x for x in input("Enter start & End Date: ex. yyyy-mm-dd|yyyy-mm-dd ").split('|')]
#             resVal = getIn()
#             num,name1 =[x for x in resVal.split('|')]
#             strDate= datetime.datetime.strptime(num, format).date()
#             endDate = datetime.datetime.strptime(name1, format).date()
#         except ValueError:
#                 msg.showerror("Formate Error",'Enter correct date format..')
#                 # quit()

#         # strDate = '2022-02-16'
#         # endDate = '2022-02-17'

#         mUrl = shekou_terminals.getUrl(strDate,endDate,pageN)
#         r = urllib.request.urlopen(mUrl)
#         output = r.read()
#         my_js = json.loads(output)

#         pageCount = my_js['TotalPages']
#         for i in range(pageCount):
#             if i ==0:
#                 df = pd.DataFrame.from_dict(my_js["InnerList"])
#                 Ndf =Ndf.append(df)

#             else:
#                 mUrl = shekou_terminals.getUrl(strDate,endDate,i)
#                 r = urllib.request.urlopen(mUrl)
#                 output = r.read()
#                 my_js = json.loads(output)
                
#                 df = pd.DataFrame.from_dict(my_js["InnerList"])
#                 Ndf =Ndf.append(df)

#         Ndf.to_excel("Terminals_Data.xlsx",index=False)
 
