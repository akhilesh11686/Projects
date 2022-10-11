#%%
from ast import Continue
from cmath import e
import pandas as pd
from tkinter import ttk,filedialog
from tkinter.filedialog import askopenfile
import os
from tkinter.messagebox import showerror
from datetime import date,datetime


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

filePath = browse('Missing VGM EDI')
df = pd.ExcelFile(filePath)
VGM_Cont = pd.ExcelFile(browse('VGM contact'))

edf = df.parse('page')
excl_POL = VGM_Cont.parse('Excluded POL of Europe')
vgmCnt = VGM_Cont.parse('VGM CONTACTS')

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
kk = rslt[(rslt['Booking Status'].astype('str').str.contains("30|60|70",regex=True,na=True)) & ~((rslt['Channel'].str.contains("0004581947|VGM_CSV channel")) & (rslt['SB_CONTACT'].astype('str').str.contains("ssc.vgm@cma-cgm.com")))]
rslt = rslt[~(rslt['Booking'].str.contains("|".join(kk['Booking'].to_list())))]

#%%

import pandas as pd
import re

rslt = pd.read_excel('s.xlsx')
#%%

print('Done')




