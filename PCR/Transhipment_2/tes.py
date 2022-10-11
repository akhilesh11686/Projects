
#%%
import pandas as pd
import win32com.client as win32
import win32com.client
import re

import tkinter as tk
from tkinter import messagebox
from openpyxl import workbook
from openpyxl import load_workbook

from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile
from PIL import Image,ImageTk
import os

root = Tk()
root.geometry('650x200')
root.resizable(False,False)
root.title('Missing VGM mails Distribution..')


def open_file(txt):
    global RwFilePath,mis_tbl,vyg,file
    
    file = askopenfile(mode='r',filetypes=[('Excel Files','*.XLSX')],title=txt)
    RwFilePath = file.name
    return RwFilePath


def xlMail_dist():

    
    df = pd.read_excel(open_file("LCR File"),sheet_name='Sheet1')
    rfMail = pd.read_excel(open_file("RFI Matrix"),sheet_name='AMS + VGM + POL QUERY')


    # ptn 1
    df1= (df[(df['Booking Status']!=1) & (df['Booking Status']!=9)])

    for i in df1['Container Number']:
        if len(df1[df1['Container Number']==i])>1:
            df1.loc[df1['Container Number']==i,'Verify Gross Mass']= max(df1[df1['Container Number']==i]['Verify Gross Mass']) 

    #ptn 2
    df2 = df1[(df1['Verify Gross Mass'].isna())]


    for id,rw in rfMail.iterrows():
        if type(rw['VGM'])!=float:
            vgmVal =rw['VGM']
        else:
            vgmVal =str(";")
            
        if type(rw['POL QUERY'])!=float:
            vgmPOL =rw['POL QUERY']
        else:
            vgmPOL=str(";")

        rfMail.loc[id,'New'] = vgmVal +";"+ vgmPOL


    Uniqu_First = df2['First POL'].drop_duplicates()

    frm = input_btn.get()

    outlook = win32com.client.Dispatch("Outlook.Application")
    oacctuse = None
    # for oac in outlook.Session.Accounts._dispobj_:
    for oac in outlook.Session.Accounts:
        if oac.DisplayName == frm:
            oacctuse = oac
            break


    
    mail = outlook.CreateItem(0)


    for i in Uniqu_First:
        # frm = 'fromMail'
        mails= rfMail[rfMail['PORTS'].str.contains(i,regex=True,na=False)]['New']
        mlBody = df2[df2['First POL']==i][['Vessel Name @ PTS','Voyage @ PTS','ETA @ PTS (GMT)','Container Number','Operator','ISO Code','CT','Teus','Verify Gross Mass','First POL','POL Code','PTS Code','Next POD','Final POD','REMARKS','Booking number']]
        subjMail = "MISSING VGM -- "


        mail = outlook.CreateItem(0)
        if oacctuse:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctuse))
        
        mail.To = ' '.join([str(elem) for elem in mails.to_list()])
        mail.Subject = subjMail  + " " +  "".join( mlBody['First POL'][:1])
        mail.HTMLBody = "Dear team,<br><br>Good day,<br><br>Please note VGM missing in LARA for below containers.<br><br>Kindly update the VGM in LARA.<br><br>{}<br><br>Note : Any charges received at transshipment port due to missing VGM, will be raised back to POL account.<br><br>".format(mlBody.to_html(header=True,index=False,justify='left',border='5'))
        # df.loc[index,'Mail_Status'] = "Sent"
        df2.loc[df2['First POL']==i,"Mail_Remark"] = "Sent"
        # mail.Display()                 
        mail.Send()        
    df2.to_excel("Sent_Mails_Status.xlsx")
    messagebox.showinfo("Thank you!!","Completed..")    
    
    
canvas = Canvas(width=550, height=230, bg='blue')
canvas.pack(expand=NO, fill=X)

image = ImageTk.PhotoImage(file="mLogo.jpg")
canvas.create_image(20, 20, image=image, anchor=NW)



pthLbl = Label(root,text=os.getlogin())
pthLbl.place(x=550,y=5)

Entr_label = tk.Label(root, text = 'Enter Generic Ids..', font = ('calibre',10,'bold'))
Entr_label.place(x=465, y=40)

input_btn = tk.Entry(root)
input_btn.place(x=465, y=70)

btn_Mail = tk.Button(root, text ='Mails_Distribution', command = lambda:xlMail_dist(),height=3,width=20)

btn_Mail.place(x=465, y=120)

root.mainloop()    


#%%
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
oacctuse = None
# #%%
# for oac in outlook.Session.Accounts._dispobj_:
#     if oac.DisplayName == 'SSC.bayplan@cma-cgm.com':
#         oacctuse = oac
#         break
