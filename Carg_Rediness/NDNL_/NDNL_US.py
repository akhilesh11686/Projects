#%%

# ivu-select-dropdown-list
import pandas as pd
from tkinter import filedialog
from datetime import date,datetime
import win32com.client
import win32com.client as win32
import os
from tkinter import messagebox
from tkinter import *   
import tkinter.font as font
from pandas.errors import EmptyDataError
import xlrd

gui = Tk(className='Cancel Booking.')
gui.geometry("500x200")


def browse(ttl):
    pth = filedialog.askopenfile(title=ttl)
    return pth

def browse1(ttl):
    pth = filedialog.askopenfilenames(title=ttl)
    return pth

def proc():
    fVsl = browse('vessel Schedule')
    Excp = browse('Exception_File')


    df = pd.read_excel(fVsl.name,skiprows=2)
    dfExc = pd.read_excel(Excp.name)

    # Blank row
    dfNBlnk= df.loc[~df['Port Code'].isnull(),['Port Code','Port Name','Voyage Ref']]

    # PortCode with PortName
    portCode = dfNBlnk[~dfNBlnk['Port Name'].isnull()]
    # portCode.to_excel('PortCode.xlsx',index=False)
    cDate = date.today()

    df = df[(~df['Voyage Ref'].isnull()) & (~df['Voyage Ref'].astype(str).str.endswith('PL'))]

    df['Date'] = pd.to_datetime(df['Cutoff']).dt.date
    df['Date'] = df['Date'].astype(str)
    CData = df[df['Date'].str.contains(str(cDate))]

    for i,rw in dfExc.iterrows():        
        if pd.isnull(dfExc.loc[i,'Service Code']):
            prtName = dfExc.loc[i,'Port Name']
            CData=CData[~CData['Port Name'].str.contains(prtName)]
        else:
            prtName = dfExc.loc[i,'Port Name']
            sCode = dfExc.loc[i,'Service Code']
            CDataN = CData.loc[(CData['Port Name']==prtName) & (CData['Service Code']!=sCode) ,CData.columns.to_list()]
            CData =CData.loc[(CData['Service Code']!=sCode) & (CData['Port Name']!=prtName),CData.columns.to_list()]        
            if len(CDataN)>0:
                CData = pd.concat([CData,CDataN],axis='rows')          

    pmrg = pd.merge(left=CData,right=portCode[['Port Code','Port Name']],left_on='Port Name',right_on='Port Name',how='left',indicator=True)
    pmrg.to_excel('Out.xlsx',index=False)
    messagebox.showinfo("Vessel_voyage created..","Done") 
def fnl():
    try:
        # voyage cleaning and mailing part heree...

        outF = pd.read_excel('Out.xlsx')

        fVyg = browse1('vygBasedFile')

        for fl in fVyg:            
            df = pd.read_csv(fl,skiprows=2,index_col=False)

            df = df[df['ASSIGNED CONTS'].astype(str).str.contains("0")]
            vyg = df['VOYAGE'].drop_duplicates().to_string(index=False)
            pol = outF.loc[outF['Voyage Ref']==vyg,'Port Code_y'].to_string(index=False)
            cDate = date.today()
            if len(df)>0:
                df = df[(df['SIZE_TYPE'].astype(str).str.contains('HC|ST|TK')) & (~df['MOVE_TERMS'].astype(str).str.contains('DP|DD|DR'))]    
                df = df[df['POL'].astype(str).str.contains(pol)]
                bkg = df['BOOKING'].drop_duplicates(keep=False)
                outF.loc[(outF['Port Code_y']==pol) &(outF['Voyage Ref']==vyg), 'Status'] = 'Sent'
                dk = pd.DataFrame({'Booking number':bkg,'Static Reason\n(Select From Dropdown or copy and paste the items in the dropdown)':'CUSTOMER:CARGO NOT READY/NO SHOW','Dynamic Reason\n(Free text with upto 30 characters)':'No activity'})
                nm = pol + " "+ vyg + " " + str(cDate.strftime('%d-%m-%y'))
                dk.to_excel(nm +'.xlsx',index=False)

                oacctuse = None

                # frm = 'ssc.vgm@cma-cgm.com'
                frm = 'ssc.usbkgcan@cma-cgm.com'

                outlook = win32com.client.Dispatch("Outlook.Application")

                for oac in outlook.Session.Accounts._dispobj_:    #<<<
                # for oac in outlook.Session.Accounts:
                    if oac.DisplayName == frm:
                        oacctuse = oac
                        break

                mail = outlook.CreateItem(0)
                if oacctuse:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctuse))

                mail.To = 'booking-cancelations@cma-cgm.com'
                mail.Subject = 'BKG CANCEL'
                mail.HTMLBody = ""
                mail.Attachments.Add(os.path.join(os.getcwd(),nm +'.xlsx'))
                mail.Display()                 
            # mail.Send()        
        outF.to_excel('Final_status.xlsx',index=False)
        messagebox.showinfo("Thank you!!","Completed..") 
    except Exception:
        pass

myFont = font.Font(family='Helvetica', size=20, weight='bold')

button = Button(gui, text='Get Voyage', bg='#0052cc', fg='#ffffff',command=proc)
button['font'] = myFont
button.pack()

button1 = Button(gui, text='Get Process', bg='#0052cc', fg='#ffffff',command=fnl)
button1['font'] = myFont

# add button to gui window
button1.pack()

gui.mainloop() 


