
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

    dd1 = pd.read_excel(open_file("LCR_11 File"))
    lst1 = dd1.iloc[:,0].to_list()
    rw1 = lst1.index('Data Source',0)
    df_11 = dd1.drop(index=dd1.index[:(rw1)])
    df_11.columns= df_11.iloc[0]
    
    # df_11 = pd.read_excel(open_file("LCR_11 File"),skiprows=17)
    # df_6 = pd.read_excel(open_file("LCR_06"),skiprows=17)
    rfMail = pd.read_excel(open_file("RFI Matrix"),sheet_name='AMS + VGM + POL QUERY')    

  

    lUnt = df_11[df_11['PTS Code']==df_11['Final POD']].index
    df_11.drop(lUnt,inplace=True)

    # exclude 1|9
    excl_1_9 =df_11[~df_11['Booking Status'].astype('str').str.contains('1|9')]

    #exclude empty flage : Y
    NonEmpty = excl_1_9[excl_1_9['Empty Flag']=='N']

    # empty vgm Verify Gross Mass
    missVGM = NonEmpty[NonEmpty['Verify Gross Mass'].isnull()]

    AvlVGM = NonEmpty[~NonEmpty['Verify Gross Mass'].isnull()]
    
    missVGM = missVGM[~missVGM['Container Number'].astype(str).isin(AvlVGM['Container Number'])]


    
    Uniqu_First = missVGM['First POL'].drop_duplicates()


    rfMail['VGM'].fillna(";",inplace=True)
    rfMail['POL QUERY'].fillna(";",inplace=True)
    rfMail['New'] = rfMail['VGM'].astype('str') + ";" + rfMail['POL QUERY'].astype('str')

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
        if len(mails)>0:
            mlBody = missVGM[missVGM['First POL']==i][['Vessel Name @ PTS','Voyage @ PTS','ETA @ PTS (GMT)','Container Number','Operator','ISO Code','Verify Gross Mass','First POL','POL Code','PTS Code','Next POD','Final POD','Booking number']]
            subjMail = "MISSING VGM -- "


            mail = outlook.CreateItem(0)
            if oacctuse:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctuse))
            
            mail.To = ' '.join([str(elem) for elem in mails.to_list()])
            mail.Subject = subjMail  + " " +  "".join( mlBody['First POL'][:1])
            mail.HTMLBody = "Dear team,<br><br>Good day,<br><br>Please note VGM missing in LARA for below containers.<br><br>Kindly update the VGM in LARA.<br><br>{}<br><br>Note : Any charges received at transshipment port due to missing VGM, will be raised back to POL account.<br><br>".format(mlBody.to_html(header=True,index=False,justify='left',border='5'))
            # df.loc[index,'Mail_Status'] = "Sent"
            missVGM.loc[missVGM['First POL']==i,"Mail_Remark"] = "Sent"
            # mail.Display()                 
            mail.Send()        
        else:
            missVGM.loc[missVGM['First POL']==i,"Mail_Remark"] = "Not Sent"

    missVGM.to_excel("Sent_Mails_Status_11.xlsx", index=False)
    messagebox.showinfo("Thank you!!","Completed..")    
    
# 06
def xl_06():
    
    dd = pd.read_excel(open_file("LCR_06"))
    lst = dd.iloc[:,0].to_list()
    rw = lst.index('Port',0)
    df_6 = dd.drop(index=dd.index[:(rw)])
    df_6.columns= df_6.iloc[0]
    # df_6 = pd.read_excel(open_file("LCR_06"),skiprows=17)    
    rfMail = pd.read_excel(open_file("RFI Matrix"),sheet_name='AMS + VGM + POL QUERY')    


    lUnt = df_6[df_6['Port']==df_6['First POL']].index
    df_6.drop(lUnt,inplace=True)

    lUnt_1 = df_6[df_6['Port']==df_6['Final POD']].index
    df_6.drop(lUnt_1,inplace=True)

    #exclude empty flage : Y
    NonEmpty = df_6[df_6['Empty']=='N']

    # here code add......................................

    # empty vgm Verify Gross Mass
    missVGM = NonEmpty[NonEmpty['Verified Gross Mass'].isnull()]
  
    AvlVGM = NonEmpty[~NonEmpty['Verified Gross Mass'].isnull()]
    
    missVGM = missVGM[~missVGM['Container Number'].astype(str).isin(AvlVGM['Container Number'])]

    Uniqu_First = missVGM['First POL'].drop_duplicates()


    rfMail['VGM'].fillna(";",inplace=True)
    rfMail['POL QUERY'].fillna(";",inplace=True)
    rfMail['New'] = rfMail['VGM'].astype('str') + ";" + rfMail['POL QUERY'].astype('str')

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
        if len(mails)>0:
            mlBody = missVGM[missVGM['First POL']==i][['Vessel Name @ CALL','Voyage @ CALL','Container Number','Operator','ISO Code','Verified Gross Mass','First POL','Next POD','Final POD','Booking Number']]
            subjMail = "MISSING VGM -- "


            mail = outlook.CreateItem(0)
            if oacctuse:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctuse))
            
            mail.To = ' '.join([str(elem) for elem in mails.to_list()])
            mail.Subject = subjMail  + " " +  "".join( mlBody['First POL'][:1])
            mail.HTMLBody = "Dear team,<br><br>Good day,<br><br>Please note VGM missing in LARA for below containers.<br><br>Kindly update the VGM in LARA.<br><br>{}<br><br>Note : Any charges received at transshipment port due to missing VGM, will be raised back to POL account.<br><br>".format(mlBody.to_html(header=True,index=False,justify='left',border='5'))
            # df.loc[index,'Mail_Status'] = "Sent"
            missVGM.loc[missVGM['First POL']==i,"Mail_Remark"] = "Sent"
            # mail.Display()                 
            mail.Send()
        else:
            missVGM.loc[missVGM['First POL']==i,"Mail_Remark"] = "Not Sent"

    missVGM.to_excel("Sent_Mails_Status_06.xlsx", index=False)
    messagebox.showinfo("Thank you!!","Completed..")    
        

canvas = Canvas(width=550, height=230, bg='blue')
canvas.pack(expand=NO, fill=X)

image = ImageTk.PhotoImage(file="mLogo.jpg")
canvas.create_image(20, 20, image=image, anchor=NW)



pthLbl = Label(root,text=os.getlogin())
pthLbl.place(x=550,y=5)

Entr_label = tk.Label(root, text = 'Enter Generic Id.', font = ('calibre',10,'bold'),height=2,width=20)
Entr_label.place(x=310, y=40)


input_btn = tk.Entry(root)
input_btn.place(x=310, y=80)

btn_Mail = tk.Button(root, text ='Mails_Distr_11', command = lambda:xlMail_dist(),height=3,width=20)
btn_Mail.place(x=465, y=120)

btn_Mail_6 = tk.Button(root, text ='Mails_Distr_06', command = lambda:xl_06(),height=3,width=20)
btn_Mail_6.place(x=310, y=120)

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
