#%%
from ast import Continue
from cmath import e
from tkinter import messagebox
import pandas as pd
from tkinter import ttk,filedialog
from tkinter.filedialog import askopenfile
import os
from tkinter.messagebox import showerror
from datetime import date,datetime
import re

from tkinter import *
import tkinter.font as font
from tkinter.messagebox import Message


gui = Tk(className='Python Examples - Button')
gui.geometry("500x100")


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

def proces():

    filePath = browse('Missing VGM EDI')
    df = pd.ExcelFile(filePath)
    VGM_Cont = pd.ExcelFile(browse('VGM contact'))
    AgncyFndr = pd.ExcelFile(browse('Agency Finder'))
    gblPort = pd.ExcelFile(browse('Global ports'))

    edf = df.parse('page')
    excl_POL = VGM_Cont.parse('Excluded POL of Europe')
    vgmCnt = VGM_Cont.parse('VGM CONTACTS')
    Agency_finder = AgncyFndr.parse('Sheet0')
    gblPort_sht = gblPort.parse('Global port codes')

    # pt3
    edf.drop('SOC',axis=1,inplace=True)
    edf['GEN Contact']= edf['GEN Contact'].str.lower()
    edf.drop_duplicates(inplace=True)

    # pt4
    edf = edf[~edf['POL'].str[:2].isin(excl_POL['Booking Prefix'])]

    # pt6
    oChannel = edf[edf['Channel'].str.contains('0004581947') & edf['SB_CONTACT'].str.contains('ssc.vgm@cma-cgm.com')]
    oChannel.to_excel("EDI_Rejec//Channel_Data.xlsx")

    # pt5
    exlChanel = edf[~edf['Booking'].str.contains(oChannel['Booking'].to_list()[0])]


    today = date.today()
    exlChanel['VGM Cutoff'].fillna(today,inplace=True)

    for id,rw in exlChanel.iterrows():
        try:
            if exlChanel.loc[id,'VGM Cutoff'].date():
                exlChanel.loc[id,'VGM Cutoff'] = exlChanel.loc[id,'VGM Cutoff'].date()
        except AttributeError:
            continue

    exlChanelN=exlChanel[exlChanel['VGM Cutoff']>=today]
    exlChanelN['Voyage_Right'] =  exlChanelN['Voyage'].str[-2:]
    exlChanelN['POL_N'] =  exlChanelN['POL'].str[:2]

    # pt7
    df = exlChanelN[(exlChanelN['Voyage_Right']!='PL') & (exlChanelN['POL_N'].str.contains('JP|TW|SG'))==False]

    # pt8
    k = df[(df['Voyage_Right']=='PL') & (df['POL_N']!='US')]
    k1 = df[(df['Voyage_Right']!='PL')]

    df_deleted = df[(df['Voyage_Right']=='PL') & (df['POL_N']=='US') & (df['Error'].str.contains("VGM provided is over Max gross weight authorized for this container, please check and provide valid VGM")==False)]


    df_PL_US_ERR = df[(df['Voyage_Right']=='PL') & (df['POL_N']=='US') & (df['Error'].str.contains("VGM provided is over Max gross weight authorized for this container, please check and provide valid VGM")==True)]
    df_PL_US_ERR['SPC_Contact'] = ""
    df_PL_US_ERR['SB_CONTACT'] = ""
    df_PL_US_ERR['POL_Op'] = df_PL_US_ERR['POL'].astype('str')+ "APL"

    for id,i in df_PL_US_ERR.iterrows():
        try:
            df_PL_US_ERR.loc[id,'Booking Contact'] = vgmCnt.loc[vgmCnt['Port Codes'].str.contains(df_PL_US_ERR.loc[id,'POL_Op']),"VGM Contact  "].to_list()[0]
        except Exception:
            continue

    frm = [k,k1,df_PL_US_ERR]
    rslt = pd.concat(frm)

    # pts 9
    rslt['SB_CONTACT'] = rslt['GEN Contact'].astype('str')+";"+rslt['SB_CONTACT'].astype('str')

    def find_email(text):
        email = re.findall(r'[\w\.-]+@[\w\.-]+',str(text))
        return ','.join(email)
    try:    
        rslt['SPC_Contact'] = rslt['SPC_Contact'].apply(lambda x:find_email(x))
        rslt['SB_CONTACT'] = rslt['SB_CONTACT'].apply(lambda x:find_email(x))
        rslt['GEN Contact'] = rslt['GEN Contact'].apply(lambda x:find_email(x))
        rslt['Booking Contact'] = rslt['Booking Contact'].apply(lambda x:find_email(x))
    except Exception:    
        Continue


    # pts 10
    rslt.loc[rslt['POL'].str.startswith('AU', na=False),'SPC_Contact']=''
    rslt.loc[rslt['POL'].str.startswith('AU', na=False),'SB_CONTACT']=''
    rslt.loc[rslt['POL'].str.startswith('NZ', na=False),'SPC_Contact']=''
    rslt.loc[rslt['POL'].str.startswith('NZ', na=False),'SB_CONTACT']=''

    rslt.loc[rslt['POL'].str.startswith('AU', na=False),'Booking Contact']='au.CargoReadiness@cma-cgm.com'
    rslt.loc[rslt['POL'].str.startswith('NZ', na=False),'Booking Contact']='nz.cargoreadiness@cma-cgm.com'


    # pts11
    rslt.loc[rslt['POL'].str.startswith('MY', na=False) & ~rslt['Voyage'].str.endswith('PL',na=False),'Booking Contact']='kua.ExportCS@cma-cgm.com'

    # pts12
    rslt.loc[rslt['POL'].str.startswith('GB', na=False),'Booking Contact']='Lpl.vgmcontact@cma-cgm.com'

    # pts13
    rslt['BOL number'] = rslt['BOL number'].apply(lambda x: x.replace("*#",""))


    # pts14
    rslt.drop(rslt[(rslt['Booking Status'].astype('str').str.contains("30|60|70",regex=True,na=True)) & ~((rslt['Channel'].str.contains("0004581947|VGM_CSV channel")) & (rslt['SB_CONTACT'].astype('str').str.contains("ssc.vgm@cma-cgm.com")))].index,inplace=True)

    rslt.loc[rslt['Booking Contact'].astype(str).str.contains('ssc.|SSC.'),'Booking Contact'] =""
    rslt['POL'].fillna('NA',inplace=True)

    for i,id in rslt.iterrows():    
        pol = id['POL']
        vyg = id['Voyage']
        if pol !='NA':
            eml = vgmCnt.loc[vgmCnt['Port Codes']==pol,'VGM Contact  ']
            if len(eml)>0:
                rslt.loc[i,'Booking Contact'] =eml.to_string(index=False)
            else:
                vyg1 = vyg[-2:]
                if vyg1=='NL':
                    yg1 = 'ANL'
                elif vyg1=='PL':
                    yg1 = 'APL'
                elif vyg1=='MA':
                    yg1 = 'CMA'

                eml = vgmCnt.loc[vgmCnt['Port Codes']==pol+yg1,'VGM Contact  ']
                if len(eml)>0:
                    rslt.loc[i,'Booking Contact'] =eml.to_string(index=False)

    # pts15
    # If agency finder not available then conside multiple email
    for i, id in rslt[rslt['Booking Contact'].isnull() & ~(rslt['POL'].str.contains('NA')) ].iterrows():
        city = gblPort_sht.loc[gblPort_sht['POINT_CODE']==id['POL'],'FULL_NAME'].to_string(index=False)
        city = city.strip()
        cntry = gblPort_sht.loc[gblPort_sht['POINT_CODE']==id['POL'],'COUNTRY NAME'].to_string(index=False)
        brnd = id['Voyage'][-2:].strip()
        cntry = cntry.strip()
        
        out =Agency_finder.loc[Agency_finder['Country'].str.contains(cntry.upper()) & Agency_finder['Operational function'].str.contains('VGM') & Agency_finder['City'].str.contains(city.upper())& Agency_finder['Brand/Agency network'].str.contains(brnd.upper()),'Email']
        if len(out)>0:
            rslt.loc[i,'Booking Contact'] = out.to_string(index=False)
        else:
            out =Agency_finder.loc[Agency_finder['Country'].str.contains(cntry.upper()) & Agency_finder['Operational function'].str.contains('VGM',regex=True) & Agency_finder['Brand/Agency network'].str.contains(brnd.upper()),'Email']
            if len(out)>2:
                rslt.loc[i,'Booking Contact'] = out.to_list()[0]

    # pts 16
    rslt.loc[rslt['BOL number'].str[:3]=='DXB','Booking Contact']=='dxb.vnair@cma-cgm.com;DXB.DLOGANATHAN@cma-cgm.com;dxb.nhiran@cma-cgm.com'
    print('Done')
    messagebox.showinfo('Done','Completed!!')

myFont = font.Font(family='Helvetica', size=20, weight='bold')
button = Button(gui, text='Start_001', bg='#0052cc', fg='#ffffff',command=proces)
button['font'] = myFont
button.pack()

gui.mainloop() 


# #%%

# import pandas as pd

# rslt = pd.read_excel('4.xlsx')
# CNHk = rslt[rslt['Booking'].str.startswith('CN|HK')]


