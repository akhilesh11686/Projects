
#%%
# ttod
import tkinter
from tkinter import messagebox
import pandas as pd
from openpyxl import workbook
from  openpyxl import load_workbook

from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile
import tkinter.font as font
import tkinter.messagebox 

import openpyxl
import datetime
from datetime import date
import win32com.client as win32
import win32com
import os
import re
import win32timezone

from tkinter.filedialog import askopenfile
from PIL import Image,ImageTk

root = Tk()
root.geometry('600x200')
root.title('OOG_Reminder_Process')

def open_file():        
    file = askopenfile(mode='r',filetypes=[('Excel Files','*.XLSX')])
    if file is not None:
        content = file.name
        return content

def proc1():
    df = pd.read_excel(open_file())
    df['Status'] = ''

    out_app = win32.gencache.EnsureDispatch('Outlook.Application')
    out_namespace= out_app.GetNamespace("MAPI")
    out_iter_folder = out_namespace.Folders['ssc.oog-ft@cma-cgm.com'].Folders['Sent Items']
    item_count = out_iter_folder.Items.Count

    my_list_mail = []
    my_list_Subj = []
    my_list_Date = []

    for i in out_iter_folder.Items:
        my_list_mail.append(i.EntryID)
        my_list_Subj.append(i.Subject)
        my_list_Date.append(str(i.SentOn))


    dk = pd.DataFrame({'Entry_ID':my_list_mail,'Sub_':my_list_Subj,'SentDate':my_list_Date})
    dk.reset_index(drop=True, inplace=True)


    for id,rw in df.iterrows():
        try:
            Subj = rw['Subject_Line']
            ou1 = dk[dk['Sub_'].str.contains(Subj,regex=True)]
            if len(ou1)>0:
                ou2 = ou1[ou1['Sub_'].str.contains('$APP',regex=False)]
                if len(ou2) >0:
                    out = ou2.loc[ou2['SentDate']==(max(ou2['SentDate'])),'Entry_ID']
                    lst = out.to_list()
                    mlItm = out_namespace.GetItemFromID(lst[0])
                    
                    rply = mlItm.ReplyAll()
                    
                    BodyN =  'Hello Partners,<br><br>Awaiting your approval.<br>'
                    rply.HTMLBody =  BodyN + rply.HTMLBody

                    
                    rply.Subject = 'Reminder :' + rply.Subject
                    # rply.Display()
                    rply.Send()
                    df.loc[id,'Status'] = "Sent"
        except:
            continue
    df.to_excel('Sent_Status_Result.xlsx')
    messagebox.showinfo('Mails Sent..','Completed!!')
# btn = Button(root, text = 'Send Reminder', command = proc1)
# btn.pack(side = 'left')   
button_font = font.Font(family='Helvitica', size=20)

button_submit = tkinter.Button(root,
    text="Reminder",
    bg='#45b592',
    fg='#ffffff',
    bd=0,
    font=button_font,
    height=2,
    width=15,command=proc1)
button_submit.grid(row = 2, column = 1, pady = 60, padx = 150)
 
button_submit = tkinter.Button(root,
    text="Reminder",
    bg='#45b592',
    fg='#ffffff',
    bd=0,
    font=button_font,
    height=2,
    width=15,command=proc1)
button_submit.grid(row = 2, column = 1, pady = 60, padx = 150)
 

root.mainloop()
