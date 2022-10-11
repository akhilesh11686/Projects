

#%%

import tkinter as ttk
from email import message
from tkinter import *
import tkinter as tk
import tkinter.font as font
from tkinter import messagebox
import sys

gui = Tk(className='Send Notification')
gui.geometry("500x150")

from datetime import datetime
import numpy as np
import pandas as pd
from tkinter import filedialog
import openpyxl    
import calendar
import numpy as np
import win32com.client as client
from pathlib import Path

def browse():
    fl = filedialog.askopenfile()
    return fl

def send_olMail():

    id=textExample.get(1.0, tk.END+"-1c")
    if len(id)==0:
        messagebox.showerror('Email id missing..','email?')
        pass


    outlook = client.Dispatch("Outlook.Application")
    ol_msg = outlook.CreateItem(0)
    ol_msg.To = id


    ol_msg.Subject = 'New indicators : - EQM Tracking -' + shName
    ol_msg.Body = 'Hi,\n\nPlease find the attached file as requested.\n\nThank you'


    attachment1 = shName + '.xlsx'
    src_file = Path.cwd() / attachment1
    ol_msg.Attachments.Add(str(src_file))

    # ol_msg.display()
    ol_msg.Send()
    messagebox.showinfo('mail sent..',"Thank you")

def getPro():
    global shName
    d = browse()

    xl = pd.ExcelFile(d.name)
    df = xl.parse(sheet_name='DATA',index_col=0)

    dd = df.index.to_list()
    skpRw = dd.index('YEAR')
    dk = df.drop(index=df.index[:skpRw], axis=0)
    header_row = dk.iloc[0]
    df1 = pd.DataFrame(dk.values[1:], columns=header_row)


    tday = datetime.today()
    yymm = str(tday.year) + "{:02d}".format(tday.month-1)
    df2 = df1[df1['MONTH'].astype('str').str.contains(yymm) & ~df1['AGENCY TYPE'].str.contains('Agency')]


    df2.loc[df2['AGENCY POINT CODE']=='AEAUH','AGENCY COUNTRY NAME'] = "AE - ABU DHABI"
    df2.loc[df2['AGENCY POINT CODE']=='AEDXB','AGENCY COUNTRY NAME'] = "AE - DUBAI"

    df2.loc[df2['MOVE COUNTRY'].str.contains("HR|ME",na=False),'AGENCY COUNTRY NAME'] = "HR - Croatia"
    df2.loc[df2['MOVE COUNTRY'].str.contains("RS",na=False),'AGENCY COUNTRY NAME'] = "RS - Serbia"

    tbl = pd.pivot_table(df2,index=['AGENCY TYPE','AGENCY COUNTRY NAME'],values=['TOTAL MOVES','MANUAL','INTEGRATION OF REJECTED'],aggfunc=np.sum)

    for id, i in tbl.iterrows():
        tbl.loc[id,'%'] = str(round((tbl.loc[id,'INTEGRATION OF REJECTED']+tbl.loc[id,'MANUAL'])/tbl.loc[id,'TOTAL MOVES']*100))+'%'
        
    mn = (tday.month-1)
    shName = calendar.month_name[mn]+ str(tday.year)
    tbl.to_excel(shName + '.xlsx')

    df5 = pd.read_excel(shName + '.xlsx')
    dic = {'AGENCY TYPE':'GBS','AGENCY COUNTRY NAME':'Agency','INTEGRATION OF REJECTED':'EDI Rejected','MANUAL':'Manual Volume','TOTAL MOVES':'Total Volume'}
    df5.rename(columns=dic,inplace=True)

    for id,i in df5.iterrows():
        if pd.isna(df5.loc[id,'GBS']):
            df5.loc[id,'GBS'] = df5.loc[id-1,'GBS']

    df5.to_excel(shName + '.xlsx',index=False)

    ss = openpyxl.load_workbook(shName + '.xlsx')
    tbl_sht = ss['Sheet1']
    tbl_sht.title='GBS List'
    ss.save(shName + '.xlsx')
    messagebox.showinfo('Done',"Thank you")

# canvas1 = tk.Canvas(gui, width = 500, height = 400)
# canvas1.pack()

# entry1 = tk.Entry(gui)
# canvas1.create_window(200, 140, window=entry1)
textExample=tk.Text(gui, height=1,width=40)
textExample.pack()

button = Button(gui, text='Data Processing.', bg='#0052cc', fg='#ffffff',command=getPro)
myFont = font.Font(family='Helvetica', size=20, weight='bold')
button['font'] = myFont
button.pack()

button1 = Button(gui, text='Send mail', bg='#0052cc', fg='#ffffff',command=send_olMail)
myFont = font.Font(family='Helvetica', size=20, weight='bold')
button1['font'] = myFont
button1.pack()

gui.mainloop()