# # date are same >> 
# # time are diffrence

# # priority ( p1,p2,p3
# # time gap > 


# #%%
# from ast import Return
# from cmath import nan
# from csv import excel
# import tkinter as tk
# from tkinter import LAST, ttk
# from tkinter import filedialog as fd
# from tkinter.messagebox import showinfo
# from numpy import NaN

# import pandas as pd
# import datetime
# from time import strptime


# root = tk.Tk()
# root.title('Tkinter comparison')
# root.resizable(False,False)
# root.geometry("350x150")

# Qlkdf = None

# # Standarisation of Partner output
# def PartnerVsl():
#     global PVsl
#     PVsl = pd.read_excel(select_file(),sheet_name=None)
    
#     # print(PVsl)
#     for i, shtName in PVsl.items():
#         if i in "Msc":
#             msc_format(shtName)
#         elif i in "Maersk":
#             maersk_format(shtName)
#             # pass
#         elif i in "Cosco":
#             Cosco_format(shtName)
#             # pass
#         elif i in "Hapag":
#             Hapag_format(shtName)

            

# def msc_format(mDf):
#     df = mDf
#     if len(df) !=0:
#         lst = []
#         dfN1 = pd.DataFrame()
#         vslUniq= df['VesselName'].drop_duplicates()
#         for vslCnt in vslUniq:
#             sVsl=df[df['VesselName'] ==vslCnt]
#             for ind in sVsl.index:
#                 if sVsl[0][ind] =='Port or Location':
#                     lst.append(sVsl[0][ind+2])
                    
#                 if sVsl[0][ind] =='Estimated Time of Arrival':
#                     lst.append(sVsl[0][ind+2])

#                 if sVsl[0][ind] =='Estimated Time of Departure':
#                     lst.append(sVsl[0][ind+2])        
#                     jdf = pd.DataFrame([lst])
#                     jdf['vessel'] = vslCnt
#                     lst.clear()
#                     dfN1= pd.concat([dfN1,jdf])
                    
#         dfN1.columns=['Port or Location','Estimated Time of Arrival','Estimated Time of Departure','vessel']
#         file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#         dfN1.to_excel(file_name,index=False,sheet_name='Msc')
#         file_name.save()   
#         print('done')

# #
# def Hapag_format(mDf):
#     df = mDf
#     if len(df) != 0:
#     # df = pd.read_excel('Partener_Vessels.xlsx',sheet_name='Hapag')
#         df.rename(columns=df.iloc[0], inplace = True)
#         df.drop(df.index[0], inplace = True)

#         for index,row in df.iterrows():
#             if pd.isnull(row['Voyage'])==False:
#                 vyVal = row['Voyage']
#             else:        
#                 row['Voyage'] = vyVal
                    
#         # dfN1.columns=['Port or Location','Estimated Time of Arrival','Estimated Time of Departure','vessel']
#         file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#         df.to_excel(file_name,index=False,sheet_name='Hapag')
#         file_name.save()   

# #


# def maersk_format(mDf):
#     df = mDf
#     if len(df) != 0:
#         lst = []
#         dfN1 = pd.DataFrame()
#         vslUniq= df['VesselName'].drop_duplicates()
#         for vslCnt in vslUniq:
#             sVsl=df[df['VesselName'] ==vslCnt]
#             for ind in sVsl.index:
#                 if 'Arrival' in sVsl['b'][ind]:
#                     lst.append(sVsl['a'][ind])
#                     lst.append(sVsl['c'][ind])
#                     lst.append(sVsl['d'][ind])
                    
#                 if 'Departure' in sVsl['a'][ind]:
#                     lst.append(sVsl['a'][ind])
#                     lst.append(sVsl['b'][ind])
#                     jdf = pd.DataFrame([lst])
#                     jdf['vessel'] = vslCnt
#                     lst.clear()
#                     dfN1= pd.concat([dfN1,jdf])

#         dfN1.columns=['Port or Location','Arrival Voyage','Estimated Time of Arrival','Departure Voyage','Estimated Time of Departure','vessel']        
#         file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#         dfN1.to_excel(file_name,index=False,sheet_name='Maersk')
#         file_name.save()   

# def Cosco_format(mDf):
#     df = mDf
#     if len(df) != 0:
#         lst = []
#         dfN1 = pd.DataFrame()
#         vslUniq= df['VesselName'].drop_duplicates()
#         for vslCnt in vslUniq:
#             sVsl=df[df['VesselName'] ==vslCnt]
#             lvyg = sVsl['voy'].dropna().to_list()[0]
#             sVsl['voy'] = lvyg
#             dfN1 = pd.concat([dfN1,sVsl])

#         file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#         dfN1.to_excel(file_name,index=False,sheet_name='Cosco')
#         file_name.save()
    

# # Comparison Part
# def QlkV():
#     global Qlkdf
#     Qlkdf = pd.read_excel(select_file())
#     df = pd.read_excel('Compare_Rw.xlsx',sheet_name=None)
#     for i, shName in df.items():
#         if i in "Msc":
#             MSC_compar(shName)
#             msc_Calc(shName)
#         elif i in "Maersk":
#             Maersk_compar(shName)
#             # pass
#         elif i in "Cosco":
#             cosco_compar(shName)
#             Cosco_Calc(shName)
#             # pass
#         elif i in "Hapag":
#             hepag_compar(shName)
    

# def MSC_compar(shName):
#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Msc')
#     df = shName
#     rQdf = Qlkdf.loc[:,['Vessel','Port','Voyage','Published EOSP (loc)','Published Berth (loc)','Published Unberth (loc)']]
#     for i in range(len(df)):
#         port = df.loc[i,'Port or Location']
#         vessl = df.loc[i,'vessel']

#         dd = rQdf[(rQdf['Vessel']==vessl) & (rQdf['Port']==port) & (rQdf['Voyage'].str.endswith('MA'))]
#         if len(dd)>0:
#             df.loc[i,"MSC_VS_CMA"] = "||"
#             df.loc[i,"CMA_Vessel"] = dd['Vessel'].to_list()[0]
#             df.loc[i,"CMA_Port"] = dd['Port'].to_list()[0]
#             df.loc[i,"CMA_Vyg"] = dd['Voyage'].to_list()[0]
#             df.loc[i,"CMA_Arr"] = dd['Published EOSP (loc)'].to_list()[0]
#             df.loc[i,"CMA_Birth"] = dd['Published Berth (loc)'].to_list()[0]
#             df.loc[i,"CMA_UN_Birth"] = dd['Published Unberth (loc)'].to_list()[0]    

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Msc')
#     file_name.save()   

# def hepag_compar(shName):
#     # import pandas as pd

#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Hapag')
#     df = shName
#     UNdf = pd.read_excel('Compare_Rw.xlsx',sheet_name='UN_Code')
#     # Get Un code
#     for i in range(len(df)):
#         port = df.loc[i,'Port']
#         unCode = UNdf[(UNdf['PORTs']==port)]
#         if len(unCode) >0:
#             df.loc[i,"Port"] = unCode['UN'].to_list()[0]

#     # Qlkdf = pd.read_excel('Qlk.xlsx',sheet_name='Sheet1')
#     # df = shName
#     rQdf = Qlkdf.loc[:,['Vessel','Port','Voyage','Published EOSP (loc)','Published Berth (loc)','Published Unberth (loc)']]

#     df.rename(columns={'Unnamed: 0':'Vessel'},inplace=True)
#     for i in range(len(df)):
#         port = df.loc[i,'Port']
#         vessl = df.loc[i,'Vessel']

#         dd = rQdf[(rQdf['Vessel']==vessl) & (rQdf['Port']==port) & (rQdf['Voyage'].str.endswith('MA'))]
#         if len(dd)>0:
#             df.loc[i,"Hapag_VS_CMA"] = "||"
#             df.loc[i,"CMA_Vessel"] = dd['Vessel'].to_list()[0]
#             df.loc[i,"CMA_Port"] = dd['Port'].to_list()[0]
#             df.loc[i,"CMA_Vyg"] = dd['Voyage'].to_list()[0]
#             df.loc[i,"CMA_Arr"] = dd['Published EOSP (loc)'].to_list()[0]
#             df.loc[i,"CMA_Birth"] = dd['Published Berth (loc)'].to_list()[0]
#             df.loc[i,"CMA_UN_Birth"] = dd['Published Unberth (loc)'].to_list()[0]    

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Hapag')
#     file_name.save()   




# def cosco_compar(shName):
#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Cosco')
#     df = shName
#     UNdf = pd.read_excel('Compare_Rw.xlsx',sheet_name='UN_Code')
#     # Get Un code
#     for i in range(len(df)):
#         port = df.loc[i,'protName']
#         unCode = UNdf[(UNdf['PORTs']==port)]
#         if len(unCode) >0:
#             df.loc[i,"protName"] = unCode['UN'].to_list()[0]

#     rQdf = Qlkdf.loc[:,['Vessel','Port','Voyage','Published EOSP (loc)','Published Berth (loc)','Published Unberth (loc)']]
#     for i in range(len(df)):
#         port = df.loc[i,'protName']
#         vessl = df.loc[i,'VesselName']

#         dd = rQdf[(rQdf['Vessel']==vessl) & (rQdf['Port']==port) & (rQdf['Voyage'].str.endswith('MA'))]
#         if len(dd)>0:
#             df.loc[i,"COSCO_VS_CMA"] = "||"
#             df.loc[i,"COSCO_Vessel"] = dd['Vessel'].to_list()[0]
#             df.loc[i,"COSCO_Port"] = dd['Port'].to_list()[0]
#             df.loc[i,"COSCO_Vyg"] = dd['Voyage'].to_list()[0]
#             df.loc[i,"COSCO_Arr"] = dd['Published EOSP (loc)'].to_list()[0]
#             df.loc[i,"COSCO_Birth"] = dd['Published Berth (loc)'].to_list()[0]
#             df.loc[i,"COSCO_UN_Birth"] = dd['Published Unberth (loc)'].to_list()[0]

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Cosco')
#     file_name.save()   


# def Maersk_compar(shName):
#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Maersk')
#     df = shName
#     UNdf = pd.read_excel('Compare_Rw.xlsx',sheet_name='UN_Code')
#     # Get Un code
#     for i in range(len(df)):
#         port = df.loc[i,'Port or Location']
#         unCode = UNdf[(UNdf['PORTs']==port)]
#         if  len(unCode)>0:
#             df.loc[i,"Port or Location"] = unCode['UN'].to_list()[0]

#     rQdf = Qlkdf.loc[:,['Vessel','Port','Voyage','Published EOSP (loc)','Published Berth (loc)','Published Unberth (loc)']]
#     for i in range(len(df)):
#         port = df.loc[i,'Port or Location']
#         vessl = df.loc[i,'vessel']

#         dd = rQdf[(rQdf['Vessel']==vessl) & (rQdf['Port']==port) & (rQdf['Voyage'].str.endswith('MA'))]
#         if len(dd)>0:
#             df.loc[i,"Maerks_VS_CMA"] = "||"
#             df.loc[i,"Maerks_Vessel"] = dd['Vessel'].to_list()[0]
#             df.loc[i,"Maerks_Port"] = dd['Port'].to_list()[0]
#             df.loc[i,"Maerks_Vyg"] = dd['Voyage'].to_list()[0]
#             df.loc[i,"Maerks_Arr"] = dd['Published EOSP (loc)'].to_list()[0]
#             df.loc[i,"Maerks_Birth"] = dd['Published Berth (loc)'].to_list()[0]
#             df.loc[i,"Maerks_UN_Birth"] = dd['Published Unberth (loc)'].to_list()[0]

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Maersk')
#     file_name.save()   

# #split sat 2nd apr 2022
# def CnvtDate(dt):
#     sVl = dt.split()
#     day1 = sVl[1].replace('nd','').replace('th','').replace('st','').replace('rd','')
#     Mnth1 = strptime(sVl[2],'%b').tm_mon   #apr >> 4
#     Yr1 = sVl[3]
#     hrm = sVl[4]
#     nDt = str(Yr1) +"-"+str(Mnth1) +"-"+str(day1) +" "+ str(hrm)
#     nDt2 = datetime.datetime.strptime(nDt, '%Y-%m-%d %H:%M')
#     nDt2=nDt2.strftime("%Y-%m-%d %H:%M")
#     return nDt2


# def msc_Calc(sdf):
#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Msc')
#     df = sdf
#     for i in range(len(df)):
#         if type(df.loc[i,'Estimated Time of Arrival']) !=float:
#             # print(df.loc[i,'Port or Location'])
#             arrTime = CnvtDate(df.loc[i,'Estimated Time of Arrival'])    
#         else:
#             arrTime = ''
#         # Convert both in str
#         C_arrTime = str(df.loc[i,'CMA_Arr'])

#         if type(df.loc[i,'Estimated Time of Departure']) != float:
#             # print(df.loc[i,'Port or Location'])
#             # print(i)
#             depTime = CnvtDate(df.loc[i,'Estimated Time of Departure'])
#         else:
#             depTime = ''

#         C_depTime = str(df.loc[i,'CMA_UN_Birth'])

#         if (len(arrTime) >0) & (len(depTime)>0)& (len(C_depTime)!='NaT') & ((C_arrTime)!='NaT'):
#             # Convert both in datetimeformate
#             arr1 = datetime.datetime.strptime(arrTime,'%Y-%m-%d %H:%M')
#             arr2 = datetime.datetime.strptime(C_arrTime,'%Y-%m-%d %H:%M:%S')
#             diff = arr2 - arr1
#             days, seconds = diff.days, diff.seconds
#             # df.loc[i,'Arr_Days_Diff'] = days
#             # df.loc[i,'Arr_Hours_Diff'] = days * 24 + seconds // 3600

#             if days == 0:
#                 df.loc[i,'Arr_Days_HRs_Diff'] = str(divmod(seconds, 3600)[0]) + " Hrs"
#             else:
#                 df.loc[i,'Arr_Days_HRs_Diff'] = str(days) + "Days"
#                 # df.loc[i,'Arr_Hours_Diff'] = days * 24 + seconds // 3600

#             arr1 = datetime.datetime.strptime(depTime,'%Y-%m-%d %H:%M')
#             arr2 = datetime.datetime.strptime(C_depTime,'%Y-%m-%d %H:%M:%S')
#             diff = arr2 - arr1
#             days, seconds = diff.days, diff.seconds
#             # df.loc[i,'Dep_Days_Diff'] = days
#             # df.loc[i,'Dep_Hours_Diff'] = days * 24 + seconds // 3600
#             if days == 0:
#                 df.loc[i,'Dep_Days_HRs_Diff'] = str(divmod(seconds, 3600)[0]) + " Hrs"
#             else:
#                 df.loc[i,'Dep_Days_HRs_Diff'] = str(days) + " Days"
#                 # df.loc[i,'Dep_Hours_Diff'] = days * 24 + seconds // 3600                        

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Msc')
#     file_name.save()           

#     # print('Done')
# def Cosco_Calc(sdf):
#     # df = pd.read_excel('Compare_Rw.xlsx',sheet_name='Cosco')
#     df = sdf
#     for i in range(len(df)):
#         if (type(df.loc[i,'arrDtlocAct']) ==float) & (type(df.loc[i,'arrDtlocCos']) ==float):
#             continue
#         elif type(df.loc[i,'arrDtlocAct']) !=float:
#             arrTime = df.loc[i,'arrDtlocAct']
#         else:
#             arrTime = df.loc[i,'arrDtlocCos']
#         # Convert both in str
#         C_arrTime = str(df.loc[i,'COSCO_Arr'])


#         if (type(df.loc[i,'depDtlocAct']) ==float) & (type(df.loc[i,'depDtlocCos']) ==float):
#             continue
#         elif type(df.loc[i,'depDtlocAct']) != float:
#             depTime = df.loc[i,'depDtlocAct']
#         else:
#             depTime = df.loc[i,'depDtlocCos']

#         C_depTime = str(df.loc[i,'COSCO_UN_Birth'])

#         if (len(arrTime) >0) & (len(depTime)>0)& (len(C_depTime)!='NaT') & ((C_arrTime)!='NaT'):
#             # Convert both in datetimeformate
#             arr1 = datetime.datetime.strptime(arrTime,'%Y-%m-%d %H:%M')
#             arr2 = datetime.datetime.strptime(C_arrTime,'%Y-%m-%d %H:%M:%S')
#             diff = arr2 - arr1
#             days, seconds = diff.days, diff.seconds
        
#             if days == 0:
#                 df.loc[i,'Arr_Days_HRs_Diff'] = str(divmod(seconds, 3600)[0]) + " Hrs"
#             else:
#                 df.loc[i,'Arr_Days_HRs_Diff'] = str(days) + "Days"

#             arr1 = datetime.datetime.strptime(depTime,'%Y-%m-%d %H:%M')
#             arr2 = datetime.datetime.strptime(C_depTime,'%Y-%m-%d %H:%M:%S')
#             diff = arr2 - arr1
#             days, seconds = diff.days, diff.seconds

#             if days == 0:
#                 df.loc[i,'Dep_Days_HRs_Diff'] = str(divmod(seconds, 3600)[0]) + " Hrs"
#             else:
#                 df.loc[i,'Dep_Days_HRs_Diff'] = str(days) + " Days"
#                 # df.loc[i,'Dep_Hours_Diff'] = days * 24 + seconds // 3600                        

#     file_name = pd.ExcelWriter('Compare_Rw.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
#     df.to_excel(file_name,index=False,sheet_name='Cosco')
#     file_name.save()       



# def select_file():
#     filetypes = (
#         ('Excel Files','*xlsx'),
#         ('All Files','*.*')
#     )

#     filename = fd.askopenfilename(
#         title = "Open file",
#         initialdir = "/",
#         filetypes= filetypes)

#     return filename

# QlckBtn = ttk.Button(root,text='QlickReport_compar',command=QlkV)
# QlckBtn.pack(expand=True)


# ptrn = ttk.Button(root,text='Std_format_partners',command=PartnerVsl)
# ptrn.pack(expand=True)

# root.mainloop()

# ____________________________________________________________
#%%
# COSCO new way *********
from tkinter import Button, messagebox
import pandas as pd
import datetime

from datetime import datetime,date, timedelta
import numpy as np
import tkinter as tk
from tkinter import LAST, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import tkinter.font as font



root = tk.Tk()
root.title('Tkinter comparison')
root.resizable(False,False)
root.geometry("650x150")

def red_font_negatives1(series):
    highlight = 'background-color: red;'
    default = ''
    return [highlight if e >= 5 & e <= 24 else default for e in series]

def yellow_font_negatives1(series):
    highlight = 'background-color: yellow;'
    default = ''
    return [highlight if e >= 1 else default for e in series] 


def orange_font_negatives1(series):
    highlight = 'background-color: orange;'
    default = ''
    return [highlight if e >= 3 else default for e in series] 


def green_font_negatives1(series):
    highlight = 'background-color: green;'
    default = ''
    return [highlight if (e <= 0) & (e >= -100) else default for e in series] 



def red_font_negatives(series):
    highlight = 'background-color: red;'
    default = ''
    return [highlight if e == '1-14' else default for e in series]

def yellow_font_negatives(series):
    highlight = 'background-color: yellow;'
    default = ''
    return [highlight if e == '30' else default for e in series] 


def orange_font_negatives(series):
    highlight = 'background-color: orange;'
    default = ''
    return [highlight if e == '15-29' else default for e in series] 


def Process_run():

    ptners = pd.read_excel('Partener_Vessels.xlsx',sheet_name='Cosco')

    ports_Service = pd.read_excel('Partener_Vessels.xlsx',sheet_name='Services')

    ports_C = pd.read_excel('Partener_Vessels.xlsx',sheet_name='Port_Conversion')

    Vlp = pd.merge(ptners,ports_C[['protName','Converted_Port']],on='protName',how='left')

    nVlp= Vlp[Vlp['Converted_Port'].notnull()]
    nVlp1= Vlp[~Vlp['Converted_Port'].notnull()]
    nVlp.drop('protName',axis=1,inplace=True)
    nVlp.rename(columns={'Converted_Port':'protName'},inplace=True)
    ptners = pd.concat([nVlp1,nVlp])

    DB_report = pd.read_excel('388004-SCHEDULE FROM DATAMART (1).xlsx',sheet_name='Extraction')

    DB_report1 = DB_report[(DB_report['IS_MAIN_VOYAGE']=='Y') & (DB_report['MAIN_CARRIER_NAME']=='CMA CGM')]
    DB_report1 = DB_report1[DB_report1['SERVICE_CODE'].isin(ports_Service['Service_Details'])]



    ptners1 = ptners
    # 'Move data not blnk to blank'
    ptners1.loc[(ptners1['arrDtlocAct'].isna()),'arrDtlocAct'] =ptners1.loc[(ptners1['arrDtlocCos'].notnull()),'arrDtlocCos']
    ptners1.loc[(ptners1['depDtlocAct'].isna()),'depDtlocAct'] =ptners1.loc[(ptners1['depDtlocCos'].notnull()),'depDtlocCos']

    # get only data exclud empty column
    ptners1 = ptners1[(ptners1['arrDtlocAct'].notnull())]


    ptners1['Greater than1'] = ''
    # currDate = date.today()-timedelta(days=3)
    currDate = date.today()

    for i,rw in ptners1.iterrows():
        cnvtFormat = datetime.strptime(rw['arrDtlocAct'], '%Y-%m-%d %H:%M')
        cnvtDate = cnvtFormat.date()    
        if (cnvtDate>=currDate) == True:
            ptners1.loc[i,'Greater than1'] = 'OK'

    ptners1 = ptners1.fillna(method='ffill')

    # remove duplicate
    ptners1 = ptners1.apply(lambda x: x.astype(str).str.lower()).drop_duplicates(subset=['protName', 'arrDtlocAct','depDtlocAct','vesselName'], keep='last')

    # get data having date more than today date
    ptners1 = ptners1[(ptners1['Greater than1']== 'OK')|(ptners1['Greater than1']== 'ok')]

    #Drop columns 
    ptners1.drop(['arrDtlocCos','depDtlocCos','Greater than1','VesselName','Converted_Port'],axis=1,inplace=True)
    
    for id,rw in DB_report1.iterrows():    
        try:
            outPt = ptners1.loc[ptners1['protName'].str.contains(rw['POINT_NAME'], case=False,na=False) &  ptners1['vesselName'].str.contains(rw['VESSEL_NAME'], case=False,na=False),'arrDtlocAct']
            outPt = pd.to_datetime(outPt)
            result = outPt.to_list()
            # lkVal  =datetime.strptime(rw['ETA_DATE'], '%Y-%m-%d %H:%M:%S')
            lkVal = rw['ETA_DATE'].to_pydatetime()
            res = min(result, key=lambda sub: abs(sub - lkVal))
            DB_report1.loc[id,'new'] = res
        except:
            continue

    DB_report1['Date_RangeGap']=''
    DB_report1['Date_diff']=''
    DB_report1['Date_Diff_Vessel']=''
    DB_report1['new'].fillna('miss',inplace=True)



    for id,rw in DB_report1.iterrows():

        if rw['new'] !='miss': 
            lara_date = rw['ETA_DATE'].to_pydatetime().date()
            Part_date = rw['new'].to_pydatetime().date()

        
            delta = lara_date-Part_date
            DB_report1.loc[id,'Date_Diff_Vessel'] = delta.days

            if rw['ETA_DATE'].to_pydatetime().date() > (date.today() + timedelta(days=30)):
                DB_report1.loc[id,'Date_RangeGap'] = '30'
                DB_report1.loc[id,'Date_diff']= rw['ETA_DATE'].to_pydatetime().date() - date.today()
            elif (rw['ETA_DATE'].to_pydatetime().date() <= (date.today() + timedelta(days=30))) & (rw['ETA_DATE'].to_pydatetime().date() >= (date.today() + timedelta(days=15))):
                DB_report1.loc[id,'Date_RangeGap'] = '15-29'
                DB_report1.loc[id,'Date_diff']= rw['ETA_DATE'].to_pydatetime().date() - date.today()
            elif rw['ETA_DATE'].to_pydatetime().date() <= (date.today() + timedelta(days=14)):
                DB_report1.loc[id,'Date_RangeGap'] = '1-14'
                DB_report1.loc[id,'Date_diff']= rw['ETA_DATE'].to_pydatetime().date() - date.today()



    DB_report11= DB_report1

    DB_report11['Date_RangeGap'].astype(str)
    DB_report11.loc[DB_report11['new']=='miss','Date_Diff_Vessel'] = -1111

    DB_report11 = DB_report11.style.apply(red_font_negatives,axis=0,subset=['Date_RangeGap'])\
        .apply(yellow_font_negatives,axis=0,subset=['Date_RangeGap'])\
        .apply(orange_font_negatives,axis=0,subset=['Date_RangeGap'])\
        .apply(red_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
        .apply(yellow_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
        .apply(orange_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
        .apply(green_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])

    DB_report11.to_excel('resutl_Cosco.xlsx')

    messagebox.showinfo("Done!!",'Completed..')


# def red_font_negatives1(series):
#     highlight = 'background-color: red;'
#     default = ''
#     return [highlight if e >= 5 & e <= 24 else default for e in series]

# def yellow_font_negatives1(series):
#     highlight = 'background-color: yellow;'
#     default = ''
#     return [highlight if e >= 1 else default for e in series] 


# def orange_font_negatives1(series):
#     highlight = 'background-color: orange;'
#     default = ''
#     return [highlight if e >= 3 else default for e in series] 


# def green_font_negatives1(series):
#     highlight = 'background-color: green;'
#     default = ''
#     return [highlight if (e <= 0) & (e >= -100) else default for e in series] 



# def red_font_negatives(series):
#     highlight = 'background-color: red;'
#     default = ''
#     return [highlight if e == '1-14' else default for e in series]

# def yellow_font_negatives(series):
#     highlight = 'background-color: yellow;'
#     default = ''
#     return [highlight if e == '30' else default for e in series] 


# def orange_font_negatives(series):
#     highlight = 'background-color: orange;'
#     default = ''
#     return [highlight if e == '15-29' else default for e in series] 
    
    

# DB_report11 = DB_report11.style.apply(red_font_negatives,axis=0,subset=['Date_RangeGap'])\
#     .apply(yellow_font_negatives,axis=0,subset=['Date_RangeGap'])\
#     .apply(orange_font_negatives,axis=0,subset=['Date_RangeGap'])\
#     .apply(red_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
#     .apply(yellow_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
#     .apply(orange_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])\
#     .apply(green_font_negatives1,axis=0,subset=['Date_Diff_Vessel'])

# DB_report11.to_excel('resutl_Cosco.xlsx')
# print('Done!!!')
myFont = font.Font(size=18)
btn = Button(root,text='Comparision',height=5,width=25,command=Process_run,bg='blue', fg='white')
btn['font'] = myFont
btn.grid(padx=150,pady=15)
root.mainloop()
