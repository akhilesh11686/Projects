# #%%
# import pandas as pd
# from datetime import datetime,date
# import win32com.client as win32
# from dateutil import parser
# from tkinter import *
# import tkinter as tk
# import tkinter.font as font
# from tkinter import messagebox
# import win32timezone
# import re

# # usrList = []
# # usrList = [item for item in input("Enter the user name : ").split("|")]
# # usrList = [item for item in input("Enter the user name : ").split("|")]
# usrList = input("Enter the user name with split ('|) value : ")

# lstGret = 'Dear|Hi|Hello|Good|Greet|From:|To:|CC|BCC|Sent:|Subject:|–'

# # usr= 'Ravi|Aarati|Siddhesh|Khuzema|Sunil|Aishwarya|Jessica|Venkatesh|Tanna|Viraj|Pranali'

# outlook = win32.Dispatch('Outlook.Application')
# namespace = outlook.GetNamespace("MAPI")
# pFolder = namespace.PickFolder()
# tFolder = pFolder.Items
# #%%
# from dateutil import parser
# # usrName = ['Ravi','Aarati','Siddhesh','Khuzema','Sunil','Aishwarya','Jessica','Venkatesh','Tanna','Viraj','Pranali']
# df12 = pd.DataFrame()
# for msg in tFolder:
#     try:
#         if msg.Class==43:                           #Mail type
#             if msg.SenderEmailType=='EX':
#                 msg.Display()
#                 mSubj = msg.Subject
#                 mBody = msg.Body
#                 mSent = msg.SentOn

#                 spVl = mBody.split('\n')
#                 df = pd.DataFrame({"Body":spVl})  

#                 #Filter based on user name
#                 fltrUser = df[df['Body'].astype('str').str.contains(usrList,regex=True,flags=re.IGNORECASE,na=True)]            
#                 if len(fltrUser)>0:
#                     #Filter based on Greetings
#                     flUsers = fltrUser[~fltrUser['Body'].astype('str').str.contains(lstGret,regex=True,flags=re.IGNORECASE,na=True)]
#                     if len(flUsers)>0:
#                         reslt = flUsers['Body'].to_list()[0]
#                         # tmVal = pd.to_datetime(msg.SentOn).strftime("%Y-%m-%d %H:%M:%S")
#                         tmVal = msg.SentOn
#                         subj = msg.Subject

#                         df12=df12.append({"SendOn":tmVal, "Subject":subj,"User" : reslt},ignore_index=True)
#                 else:
#                     pass                                                    
#             else:            
#                 pass
#     except Exception:
#         pass
# df12.to_excel('11.xlsx',index=False)
 #%%

import pandas as pd
import re
import sys
from tkinter import filedialog
from openpyxl import load_workbook
from tkinter import *
import tkinter.font as font
from tkinter import messagebox
from tkinter import ttk
gui = Tk(className='Consolidation_Data')
gui.geometry("500x200")


def rWCnt(df,lbl):
    strRow = df[df['Service Contract:'].astype('str').str.contains(lbl,regex=True,flags= re.IGNORECASE)].index[0]
    return strRow

def brows(file):
    pth = filedialog.askopenfile(title=file)
    return pth.name

def myFunc(e):
  return len(e)    


def pro():
    # strField = 'SPECIAL EQUIPMENT, HAZARDOUS, SOC, OOG...'
    # endField = 'ORIGIN ARBITRARY TABLE'
    strField = entry.get()
    endField = entry2.get()
    df1 = pd.ExcelFile(brows('ChooseFile'))
    k = 21
    list_clmn = []
    dfList = []
    for sht in df1.sheet_names:
        try:
            
            if "APPENDIX" in sht:
                # dSht += sht + "|"
                df = df1.parse(sht)
                # df = pd.read_excel('1.xlsx')
                # df = x
                try:
                    strRow = rWCnt(df, strField)
                    endRow = rWCnt(df, endField)
                except IndexError:
                    print('Lable missing')
                    sys.exit()
                    
                df_Tbl = df.iloc[strRow:endRow]
                lst = df_Tbl.isin(['Place of Receipt']).any(axis=1).to_list()
                indVal = lst.index(True)

                # df_Tbl.columns = df_Tbl.iloc[indVal]
                df_Tbl.drop(df_Tbl.index[:indVal],axis=0,inplace=True)
                df_Tbl.columns = df_Tbl.iloc[0]

                for i in range(0,df_Tbl.shape[0]):
                    for j in range(0,df_Tbl.shape[1]):
                        if pd.notnull(df_Tbl.iloc[(i+1),j]):
                            df_Tbl.iloc[i,j] = str(df_Tbl.iloc[i,j])+"|"+str(df_Tbl.iloc[(i+1),j])        
                    break

                df_Tbl = pd.DataFrame(df_Tbl.values[2:],columns=df_Tbl.iloc[0])
                df_Tbl['Appendix_n'] = sht

                # putting dataframe into list
                for n in df_Tbl.columns:
                    list_clmn.append(n)
                # df_Tbl.to_excel("RESULT_"+ str(k)+ ".xlsx")
                dfList.append(df_Tbl)
                k += 1
        except AttributeError:
            pass    
    # print(list_clmn)

    # remove duplicate  
    kk = list(set(list_clmn))

    # remove duplicate and space
    nn = list(set([str(x).replace(" ","") for x in kk if pd.notna(x)]))
    nn.sort(reverse=False, key=myFunc)

    df2 = pd.DataFrame(columns=nn)

    nDF = pd.DataFrame()
    for d in dfList:
        df_l = d
        df_l = df_l.fillna("-")
        # Clean the column replace space
        nlst = [str(x).replace(" ","") for x in df_l.columns]
        
        # loop of common excel file
        for cl in df2.columns:
            if  cl in nlst:
                print(cl)
                getIndx = nlst.index(cl)
                df2[cl]=df_l.iloc[:,getIndx]
        df2 = df2.fillna("-")
        nDF = pd.concat([df2,nDF])

    nDF.to_excel('result.xlsx')
    messagebox.showinfo("Completed!!","Thank you..")

myFont = font.Font(family='Helvetica', size=12, weight='bold')

lbl = ttk.Label(gui, text='From Table')
lbl['font'] = myFont
lbl.place(relx=0.0, rely=0.1, anchor='w')

entry = Entry(gui, width= 42)
entry.place(relx= 0.5, rely= 0.1, anchor= CENTER)


lbl = ttk.Label(gui, text='End Table')
lbl['font'] = myFont
lbl.place(relx=0.0, rely=0.3, anchor='w')

entry2 = Entry(gui, width= 42)
entry2.place(relx= 0.5, rely= 0.3, anchor= CENTER)

button = Button(gui, text='Extract Table', bg='#0052cc', fg='#ffffff',height= 2, width=10,command=pro)
button['font'] = myFont
button.place(relx=0.5, rely=0.8, anchor=CENTER)


gui.mainloop() 


#%%











