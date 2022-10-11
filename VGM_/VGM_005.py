
#%%

from turtle import title
import pandas as pd 
# pip install pretty-html-table
from pretty_html_table import build_table
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import showerror

from tkinter.filedialog import askopenfile
# import os
import re
import win32com.client as win32
import os
# import pandas as pd
from pathlib import Path

from tkinter import *
import tkinter.font as font

gui = Tk(className='VGM process')
gui.geometry("500x300")


def browse(txt):
    global sPath
    file_path = filedialog.askopenfilename(initialdir = os.getcwd(),title= txt,filetypes = (("Excel file","*.xl*"),("All files","*.*")))

    if os.path.isfile(file_path):
        try:
            file_path.split('.')[-1] == 'xlsx'
            sPath = file_path
            # df = pd.read_excel(sPath, sheet_name='page')
            return sPath 
            
        except Exception as e:
            showerror("Incorrect File", "had an issue opening file")
    else:
        showerror("File not selected", "File missing!")


def process():

    filePath = browse('Missing VGM')
    df = pd.ExcelFile(filePath)
    VGM_Cont = pd.ExcelFile(browse('VGM contact'))
    GBL_Port = pd.ExcelFile(browse('Global Ports'))
    Agency_fnd = pd.ExcelFile(browse('Agency_finder'))
    Booker_ID = pd.ExcelFile(browse('Booker_Emails'))

    Bkr = Booker_ID.parse('Booker')
    df_data = df.parse('page')
    df_data['deleted_Lanes'] = ''
    # re.findall(r'(.*)day','-1 day 5 hours 39 minutes 47 seconds')
    for id, i in df_data.iterrows():
        out = re.findall(r'(.*)day',i['Remaining time before cut off'])
        try:
            if int(out[0]) <-3:
                df_data.loc[id,'deleted_Lanes'] = 'Del'
        except IndexError:
            continue
        
    df_data.drop(df_data[df_data['deleted_Lanes']=='Del'].index,inplace=True)
    df_data.drop('deleted_Lanes',axis=1,inplace=True)


    #2 exclude 
    VGM_ContE = VGM_Cont.parse('Excluded BKG prefix - CHASER',skiprows=2)
    df_data['Job_Bkg_POL'] = df_data['Job Reference'].str[:3] + df_data['Booking POL'].str[:2]


    total_merg = pd.merge(df_data,VGM_ContE[['BOOKING_AGENT']],how='outer',left_on='Job_Bkg_POL',right_on='BOOKING_AGENT',indicator=True)
    df_1 = total_merg[total_merg['_merge']=='left_only']

    df_data.drop('Job_Bkg_POL',axis=1,inplace=True)
    df_1['Job_Bkg_POL'] = df_1['Job_Bkg_POL'].str[:3]
    df_1.drop('_merge',axis=1,inplace=True)


    #3 Bkg prefix list 
    VGM_Cont_Prefix_list = VGM_Cont.parse(' Bkg Prefix list')
    vlFnd = filePath.lower().find('europe')
    if vlFnd!= -1:
        VGM_Cont_Prefix_list = VGM_Cont_Prefix_list.iloc[17:26,0:5]
    else:
        VGM_Cont_Prefix_list = VGM_Cont_Prefix_list.iloc[0:16,0:5]
        VGM_Cont_Prefix_list['Carrier'].replace('CMA CGM','CMA',inplace=True)
        VGM_Cont_Prefix_list['Keys']=VGM_Cont_Prefix_list['Booking Prefix']+VGM_Cont_Prefix_list['Carrier']

    df_1['keys'] = df_1['Job_Bkg_POL']+df_1['Operator']
    total_merg = pd.merge(df_1,VGM_Cont_Prefix_list[['Keys','VGM contact']],left_on='keys',right_on='Keys',how='outer',indicator=True)
    df_2 = total_merg[total_merg['_merge']!='right_only']
    df_2.drop('Job_Bkg_POL',axis=1,inplace=True)
    df_2.loc[df_2['_merge']=='left_only','Operator_'] =  df_2['Booking POL'].astype(str)+df_2['Operator'].astype(str)
    df_2.drop(['keys','Keys'],axis=1,inplace=True)


    #4 VGM contact
    VGM_Cont_C = VGM_Cont.parse('VGM CONTACTS')

    VGM_Cont_C.loc[VGM_Cont_C['Port Codes'].str.len()==5,'Port Codes'] = VGM_Cont_C['Port Codes'].astype(str)+"CMA"

    df_2.drop('_merge',axis=1,inplace=True)
    total_merg = pd.merge(df_2,VGM_Cont_C[['Port Codes','VGM Contact  ']],how='outer',left_on='Operator_',right_on='Port Codes',indicator=True)
    df_3 = total_merg[total_merg['_merge']!='right_only']
    df_3.drop('_merge',axis=1,inplace=True)


    #5 Global port code

    GBL_Port_code= GBL_Port.parse('Global port codes')
    total_merg = pd.merge(df_3,GBL_Port_code[['POINT_CODE','COUNTRY NAME','FULL_NAME']],how='outer',left_on='Booking POL',right_on='POINT_CODE',indicator=True)
    df_4 = total_merg[total_merg['_merge']!='right_only']
    df_4.drop('_merge',axis=1,inplace=True)


    Agency_fndr = Agency_fnd.parse('Sheet0')
    Agency_fndr['Brand/Agency network']= Agency_fndr['Brand/Agency network'].str.replace('CMACGM','CMA')
    # Agency_fndr['Brand/Agency network'] = Agency_fndr['Brand/Agency network'].str.lower()

    Agency_fndr = Agency_fndr[['Point of contact name','Country','City','Brand/Agency network','Operational function','Email']]
    Agency_fndr_VGM = Agency_fndr[Agency_fndr['Operational function'].str.contains('VGM Contact',regex=True,na=True)]
    Agency_fndr_CR_CC = Agency_fndr[Agency_fndr['Operational function'].str.contains('Customer Service|Cargo readiness',regex=True,na=True)]


    # Replace space from string
    df_4['COUNTRY NAME'] = df_4['COUNTRY NAME'].str.strip()
    df_4['COUNTRY NAME'] = df_4['COUNTRY NAME'].str.upper()
    df_4['FULL_NAME'] = df_4['FULL_NAME'].str.strip()
    df_4['FULL_NAME'] = df_4['FULL_NAME'].str.upper()

    Agency_fndr_VGM['Brand/Agency network'].fillna("Missing",inplace=True)



    for id,rw in df_4.iterrows():  
        try:
            outVgm = Agency_fndr_VGM.loc[Agency_fndr_VGM['Country'].str.contains(rw['COUNTRY NAME'])& (Agency_fndr_VGM['City'].str.contains(rw['FULL_NAME']))& (Agency_fndr_VGM['Brand/Agency network'].str.contains(rw['Operator'])),'Email']
            if len(outVgm)==0:
            # Without Brand
                outVgm = Agency_fndr_VGM.loc[Agency_fndr_VGM['Country'].str.contains(rw['COUNTRY NAME'])& (Agency_fndr_VGM['City'].str.contains(rw['FULL_NAME']))& (Agency_fndr_VGM['Brand/Agency network'].str.contains("Missing")),'Email']

            if len(outVgm)>0:            
                output1 = ';'.join(outVgm)
                df_4.loc[id,'Agency_Email'] = output1
            else:
                outV = Agency_fndr_CR_CC.loc[Agency_fndr_CR_CC['Country'].str.contains(rw['COUNTRY NAME'])& (Agency_fndr_CR_CC['City'].str.contains(rw['FULL_NAME']))& (Agency_fndr_CR_CC['Brand/Agency network'].str.contains(rw['Operator'])),'Email']
                if len(outV)==0:
                    # Without Brand                
                    outV = Agency_fndr_CR_CC.loc[Agency_fndr_CR_CC['Country'].str.contains(rw['COUNTRY NAME'])& (Agency_fndr_CR_CC['City'].str.contains(rw['FULL_NAME']))& (Agency_fndr_CR_CC['Brand/Agency network'].str.contains("Missing")),'Email']

                output = ';'.join(outV)
                df_4.loc[id,'Agency_Email'] = output
        except Exception as e:
            continue

    # vgm contact remove agency name
    df_4.loc[~((df_4['VGM contact'].isna()==True) & (df_4['VGM Contact  '].isna()==True)),'Agency_Email'] = ''

    df_4['VGM contact'] = df_4['VGM contact'].fillna(";")
    df_4['VGM Contact  '] = df_4['VGM Contact  '].fillna(";")
    df_4['Agency_Email'] = df_4['Agency_Email'].fillna(";")

    # df_4[(df_4['VGM contact'].isna()!=True) & (df_4['VGM Contact  '].isna()!=True)]

    df_4['Combo_Emails'] = df_4[['VGM contact','VGM Contact  ','Agency_Email']].apply(lambda x:';'.join(x.astype(str)),axis=1)

    newdf = df_4.drop_duplicates(
    subset = ["Booking POL","PTS","Booking POD","Pool Location","Service","Voyage","Vessel Name","ETA Local Date @ Port","ETB Local Date @ Port","ETD Local Date @ Port","Operator","Job Reference","Part Load Booking","Mother Booking Ref.","Container Number","Commodity","Package","SOC Flag","E-Booking Requester","GEN Email(s)","BOC Email(s)","Booker Id","Booker Id Email","Cut Off Local Date","Cut Off Indian Date","Remaining time before cut off","Verified Gross Mass Unit","Verified Gross Mass","Booking/BL Gross Weight","Shipper Name"],
    keep = 'first').reset_index(drop = True)

    newdf = newdf.drop(['BOOKING_AGENT','VGM contact','Operator_','Port Codes','VGM Contact  ','POINT_CODE','COUNTRY NAME','FULL_NAME','Agency_Email'], axis=1)

    total_merg = pd.merge(left=newdf,right=Bkr,left_on='Booker Id',right_on='Name',how='outer',indicator=True)
    df_5 = total_merg[total_merg['_merge']!='right_only']

    df_5.loc[(df_5['Combo_Emails']==';;;;') & (df_5['_merge']=='both'),'Combo_Emails']=df_5['Email_Id']

    df_5.drop(['Name','Email_Id','_merge'],axis=1,inplace=True)


    df_5.loc[(df_5['Remaining time before cut off'].str.contains('(.*)day',regex=True)) & (df_5['Remaining time before cut off'].str.contains('-1',regex=True)),'Chaser_Type'] = 'Chaser2'
    df_5.loc[(df_5['Remaining time before cut off'].str.contains('(.*) hours|(.*) hour',regex=True)) ,'Chaser_Type'] = 'Chaser2'
    df_5.loc[(df_5['Remaining time before cut off'].str.contains('(.*)day',regex=True)) & (df_5['Remaining time before cut off'].str.contains('-2|-3',regex=True)),'Chaser_Type'] = 'Chaser1'
    df_5.loc[(df_5['Chaser_Type']!="Chaser1") & (df_5['Chaser_Type']!="Chaser2"),'Chaser_Type'] = 'Chaser3'



    df_5.loc[~(df_5['Combo_Emails'].str.contains('@')==True) & (df_5['E-Booking Requester'].str.contains('@')==True),'Combo_Emails'] = df_5['E-Booking Requester']
    df_5.loc[~(df_5['Combo_Emails'].str.contains('@')==True) & ~(df_5['E-Booking Requester'].str.contains('@')==True) & (df_5['GEN Email(s)'].str.contains('@')==True),'Combo_Emails'] = df_5['GEN Email(s)']
    df_5.loc[~(df_5['Combo_Emails'].str.contains('@')==True) & ~(df_5['E-Booking Requester'].str.contains('@')==True) & ~(df_5['GEN Email(s)'].str.contains('@')==True) & (df_5['BOC Email(s)'].str.contains('@')==True),'Combo_Emails'] = df_5['BOC Email(s)']
    df_5.loc[~(df_5['Combo_Emails'].str.contains('@')==True) & ~(df_5['E-Booking Requester'].str.contains('@')==True) & ~(df_5['GEN Email(s)'].str.contains('@')==True) & ~(df_5['BOC Email(s)'].str.contains('@')==True) & (df_5['Booker Id Email'].str.contains('@')==True) ,'Combo_Emails'] = df_5['Booker Id Email']

    df_5.to_excel('ResulF.xlsx')

# ####################################################################### Mails start here #################################
# import win32com.client as win32
# import os
# import pandas as pd
# from pathlib import Path

def chase1():

    bkID = pd.ExcelFile('Bookler_Emails_List.xlsx')
    mailT = bkID.parse('m5')
    df_5 = pd.read_excel('ResulF.xlsx')

    df_5_1= df_5[df_5['Chaser_Type']=='Chaser1']
    if len(df_5_1)>1:
        UnqJobRef = df_5_1['Job Reference'].drop_duplicates()

        mailFile=os.curdir + "\\VGMProcessTemplate.xlsx"
        outlook = win32.Dispatch('outlook.application')
        oacctouse = None
        for oacc in outlook.Session.Accounts:
        # for oacc in outlook.Session.Accounts._dispobj_:
            if oacc.SmtpAddress == 'ssc.vgm@cma-cgm.com':
                oacctouse = oacc
                break

        # Loop for job reference
        for rData in UnqJobRef:
            mail = outlook.CreateItem(0)
            if oacctouse:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))

            tbl = df_5_1[(df_5_1['Job Reference']== rData)]
            nTbl = tbl[['Booking POL','Voyage','Vessel Name','Job Reference','Container Number','Cut Off Local Date','Remaining time before cut off']]
            nTbl.reset_index().rename(columns={'Job Reference':'Booking','Container Number':'Container','Booking POL':'POL','Voyage':'Voyage Ref','Cut Off Local Date':'VGM Cut-off','Remaining time before cut off':'Day remaining to send VGM'},inplace=True)
            nTbl.reset_index(inplace=True)
            nTbl = nTbl.drop(['index'],axis=1)
            mailT['Body11'].fillna("",inplace=True)

            mail.To = tbl['Combo_Emails'].to_list()[0]
            mail.Subject = 'VGM declaration Missing'
            # mail.attachement = mailFile
            kk = str(Path().absolute()) + "\\VGMProcessTemplate.xlsx"
            mail.Attachments.Add(kk)
            
            # nTbl.drop(['index'],axis=1)
            mail.HTMLBody = "{0}".format(mailT.to_html(header=False,index=False,justify='left',border='0'))
            mail.HTMLBody = mail.HTMLBody.replace("{0}",nTbl.to_html())
            # mail.Display()
            df_5_1.loc[(df_5_1['Job Reference']== rData)&(df_5_1['Chaser_Type']=='Chaser1'),'sent_status'] = 'Sent'
            mail.Send()
    df_5_1.to_excel('Result_Chaser1.xlsx')


# import win32com.client as win32
# import os
# import pandas as pd
# from pathlib import Path

def chase2():
    bkID = pd.ExcelFile('Bookler_Emails_List.xlsx')
    mailT = bkID.parse('m5_2')
    df_5 = pd.read_excel('ResulF.xlsx')

    df_5_2= df_5[df_5['Chaser_Type']=='Chaser2']
    if len(df_5_2)>1:


        report_path = 'Excel'
        if not os.path.exists(report_path):
            os.makedirs(report_path)    

        UnqJobRef = df_5_2['Voyage'].drop_duplicates()

        mailFile=os.curdir + "\\VGMProcessTemplate.xlsx"
        outlook = win32.Dispatch('outlook.application')
        oacctouse = None
        # for oacc in outlook.Session.Accounts._dispobj_:
        for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == 'ssc.vgm@cma-cgm.com':
                oacctouse = oacc
                break

        # Loop for job reference
        for rData in UnqJobRef:
            mail = outlook.CreateItem(0)
            if oacctouse:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))

            tbl = df_5_2[(df_5_2['Voyage']== rData)]
            nTbl = tbl[['Booking POL','Voyage','Vessel Name','Job Reference','Container Number','Cut Off Local Date','Remaining time before cut off','Booker Id Email']]
            nTbl.rename(columns={'Job Reference':'Booking','Container Number':'Container','Booking POL':'POL','Voyage':'Voyage Ref','Cut Off Local Date':'VGM Cut-off','Remaining time before cut off':'Day remaining to send VGM','Booker Id Email':'Customer email addresses'},inplace=True)
            nTbl.reset_index(inplace=True)
            nTbl = nTbl.drop(['index'],axis=1)
            nTbl.to_csv(str(Path().absolute()) + "\\Excel\\"+'chaser2.csv')

            mailT['Body11'].fillna("",inplace=True)

            mail.To = tbl['Combo_Emails'].to_list()[0]

            vyg = tbl['Voyage'].to_list()[0]
            POL = tbl['Booking POL'].to_list()[0]
            cutoff = str(tbl['Cut Off Local Date'].to_list()[0])

            mail.Subject = 'VGM CHASER SUMMARY -'+ vyg + " - " + POL + " - " + 'VGM Cut-off : ' + cutoff
            # mail.attachement = mailFile
            kk = str(Path().absolute()) + "\\VGMProcessTemplate.xlsx"
            kk2 = str(Path().absolute()) + "\\Excel\\"+'chaser2.csv'

            mail.Attachments.Add(kk)
            mail.Attachments.Add(kk2)
            # nTbl.drop(['index'],axis=1)
            mail.HTMLBody = "{0}".format(mailT.to_html(header=False,index=False,justify='left',border='0'))
            mail.HTMLBody = mail.HTMLBody.replace("{0}",nTbl.to_html())
            # mail.Display()
            df_5_2.loc[(df_5_2['Voyage']== rData)&(df_5_2['Chaser_Type']=='Chaser1'),'sent_status'] = 'Sent'
            mail.Send()
            # remove file 
            os.remove(kk2)
    df_5_2.to_excel('Result_chaser2.xlsx')



# import win32com.client as win32
# import os
# import pandas as pd
# from pathlib import Path

def chase3():
    bkID = pd.ExcelFile('Bookler_Emails_List.xlsx')
    mailT = bkID.parse('m5_3')
    df_5 = pd.read_excel('ResulF.xlsx')

    df_5_3= df_5[df_5['Chaser_Type']=='Chaser3']
    if len(df_5_3)>1:
        UnqJobRef = df_5_3['Voyage'].drop_duplicates()

        mailFile=os.curdir + "\\VGMProcessTemplate.xlsx"
        outlook = win32.Dispatch('outlook.application')
        oacctouse = None
        # for oacc in outlook.Session.Accounts._dispobj_:
        for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == 'ssc.vgm@cma-cgm.com':
                oacctouse = oacc
                break

        # Loop for job reference
        for rData in UnqJobRef:
            mail = outlook.CreateItem(0)
            if oacctouse:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))

            tbl = df_5_3[(df_5_3['Voyage']== rData)]
            nTbl = tbl[['Booking POL','Voyage','Vessel Name','Job Reference','Container Number','Cut Off Local Date','Remaining time before cut off']]
            nTbl.reset_index().rename(columns={'Job Reference':'Booking','Container Number':'Container','Booking POL':'POL','Voyage':'Voyage Ref','Cut Off Local Date':'VGM Cut-off','Remaining time before cut off':'Day remaining to send VGM'},inplace=True)
            nTbl.reset_index(inplace=True)
            nTbl = nTbl.drop(['index'],axis=1)
            mailT['Body11'].fillna("",inplace=True)

            mail.To = tbl['Combo_Emails'].to_list()[0]
            mail.Subject = 'Attention: VGM declaration Missing post VGM Cut-off'
            # mail.attachement = mailFile
            kk = str(Path().absolute()) + "\\VGMProcessTemplate.xlsx"
            mail.Attachments.Add(kk)
            
            # nTbl.drop(['index'],axis=1)
            mail.HTMLBody = "{0}".format(mailT.to_html(header=False,index=False,justify='left',border='0'))
            mail.HTMLBody = mail.HTMLBody.replace("{0}",nTbl.to_html())
            # mail.Display()
            df_5_3.loc[(df_5_3['Voyage']== rData),'sent_status'] = 'Sent'
            mail.Send()
    df_5_3.to_excel('Result_Chaser3.xlsx')

myfont = font.Font(family='Verdana',size=10,weight='bold')
bPro = Button(gui,text='Process Data',bg='#0052cc', fg='#ffffff',command=process)
bPro['font'] = myfont
bPro.pack()

bChase1 = Button(gui,text='Chaser1',bg='#0052cc', fg='#ffffff',command=chase1)
bChase1['font'] = myfont
bChase1.pack()

bchase2 = Button(gui,text='Chaser2',bg='#0052cc', fg='#ffffff',command=chase2)
bchase2['font'] = myfont
bchase2.pack()

bchase3 = Button(gui,text='Chaser3',bg='#0052cc', fg='#ffffff',command=chase3)
bchase3['font'] = myfont
bchase3.pack()



gui.mainloop()
