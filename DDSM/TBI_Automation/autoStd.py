#%%
#7july2022
from itertools import count
from tkinter import *
from tkinter import filedialog
from xml.dom.pulldom import IGNORABLE_WHITESPACE
import pandas as pd
from tkinter import messagebox
import numpy as np

def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
        title="Select a File",
        filetypes=(("Excel Files","*.xl*"),("all Files","*.*"))
    )
        # Change label contents
    label_file_explorer.configure(text="File Opened: "+filename)

def OutProcess():
    
    df = pd.read_excel(filename,sheet_name="Sheet1")
    dfHW = pd.read_excel(filename,sheet_name="Hawaii")
    dfRO = pd.read_excel(filename,sheet_name="Rollover")

    # df = pd.read_excel("Exp Data.xlsx",sheet_name="Sheet1")
    # dfHW = pd.read_excel("Exp Data.xlsx",sheet_name="Hawaii")
    # dfRO = pd.read_excel("Exp Data.xlsx",sheet_name="Rollover")

    df = df[df['status']=="Pending TBI"]
    df_Non_pending_TBI = df[df['status']!="Pending TBI"]

    # ptn2.
    df.loc[((df['status']=="Pending TBI" ) & (~df['START_LOCATION'].str.startswith("CA") & ~df['START_LOCATION'].str.startswith("US"))),"Comment"] = "Overseas Pairing"

    # ptn3.
    df.loc[((df['status']=="Pending TBI" ) & (~df['Stop Location'].str.startswith("CA") & ~df['Stop Location'].str.startswith("US"))),"Comment"] = "Overseas Pairing"

    # ptn4.
    df.loc[((df['status']=="Pending TBI" ) & (df['START_LOCATION'].str.startswith("CA")) & (df['Stop Location'].str.startswith("CA"))),"Comment"] = "Canada Pairing"
    # df.loc[((df['status']=="Pending TBI" ) & (df['Stop Location'].str.startswith("CA"))),"Comment"] = "Canada Pairing"

    # ptn5.
    df.loc[((df['status']=="Pending TBI" ) & (df['Equipment Size/Type'].str.endswith("GA") | df['Equipment Size/Type'].str.endswith("GO"))),"Comment"] = "Genset Rejection"


    # ptn6.
    HwLst = dfHW['Hawaii Location'].to_list()
    df.loc[((df['status']=="Pending TBI" ) & (df['START_LOCATION'].isin(HwLst)) & (df['Stop Location'].isin(HwLst))),"Comment"] = "Hawaii Location"

    # ptn7.
    df.loc[((df['status']=="Pending TBI") & df['Start Status'].isin(["REX"])) & (df['Stop Status'].isin(['MEA','MED'])) & (df['Job Reference'].isnull()),"Comment"] = "Domestic booking – No Bill"

    # ptn8.
    df.loc[((df['status']=="Pending TBI") & df['Start Status'].isin(['REX','MOS'])) & (df['Stop Status'].isin(['MEA','MED'])) & (df['Comment'].isnull()),"Comment"] = "No Bill"



    # ptn10.
    dLatest = dfRO[['BKH - Booking Ref','VR - Rollover Date','VR - Fault']]
    dLatest['a_type_date'] = [min(dLatest[(dLatest['BKH - Booking Ref']==i)]['VR - Rollover Date']) for i in dLatest['BKH - Booking Ref']]
    
    # dfRO_filter1 = Filter_Latest[Filter_Latest['Job Reference'].isin(dds11.index.to_list()) & Filter_Latest['VR - Rollover Date'].isin(dds11['Validity1'].to_list())]
    Filter_Latest11 = dLatest[['BKH - Booking Ref','VR - Rollover Date','VR - Fault','a_type_date']]
    Notempty_Filter_Latest11 = Filter_Latest11[Filter_Latest11['a_type_date'].notnull()]
    Notempty_Filter_Latest11 = Notempty_Filter_Latest11[Notempty_Filter_Latest11['VR - Rollover Date'] == Notempty_Filter_Latest11['a_type_date']]
    Notempty_Filter_Latest11 = Notempty_Filter_Latest11.drop_duplicates()
    df['Foult'] = [Notempty_Filter_Latest11.loc[Notempty_Filter_Latest11['BKH - Booking Ref'] == i,'VR - Fault'].values for i in df['Job Reference']]
    df['Foult_1'] = [np.size(i) for i in df['Foult']]
    df.loc[df['Foult_1'] == 0 ,"Foult" ] = ""

    df.to_excel("OutN.xlsx")
    df= pd.read_excel('outN.xlsx')
# ################################33

    df.loc[df['Comment'].isin(['Hawaii Location']),'Comment'] = df['Foult']
    df.loc[df['Comment'].isna(),'Comment'] = df['Foult']


    for id,rw in df.iterrows():
        if df.loc[id,"Comment"] == "['CUSTOMER']": 
            df.loc[id,"Comment"] = 'Customer Rollover'
        elif df.loc[id,"Comment"] == "['CARRIER']":
            df.loc[id,"Comment"] = 'Carrier Rollover'

    Left_join = df
    Left_join['Split'] = Left_join['Job Reference'].str[-1:]

    zero_Val = dfRO[dfRO['VR - Fault'].isna()]

    lstvl = zero_Val['BKH - Booking Ref'].to_list()
    Left_join.loc[(Left_join['Foult'].isna()) & (Left_join['Job Reference'].isin(lstvl)),"Foult"] = 0

    Left_join.loc[(Left_join['Foult']==0) & Left_join['Split'].str.contains('[0-9]',regex=True),'Comment'] = 'No Rollover' 
    Left_join.loc[(Left_join['Foult']==0) & Left_join['Split'].str.contains('[A-Z]',regex=True),'Comment'] = 'No Rollover – Split' 


    #ptn12.
    Left_join['Rule 1'] = ''

    dLatest = dfRO[['BKH - Booking Ref','VR - Rollover Date','VR - Fault']]
    dLatest['a_type_date'] = [max(dLatest[(dLatest['BKH - Booking Ref']==i)]['VR - Rollover Date']) for i in dLatest['BKH - Booking Ref']]

    Filter_Latest12 = dLatest[['BKH - Booking Ref','VR - Rollover Date','VR - Fault','a_type_date']]
    Notempty_Filter_Latest12 = Filter_Latest12[Filter_Latest12['a_type_date'].notnull()]
    Notempty_Filter_Latest12 = Notempty_Filter_Latest12[Notempty_Filter_Latest12['VR - Rollover Date'] == Notempty_Filter_Latest12['a_type_date']]
    Notempty_Filter_Latest12 = Notempty_Filter_Latest12.drop_duplicates()

    Left_join.loc[(Left_join['Job Reference'].isin(lstvl)),"Modify date"] = 0
    Left_join['Modify date_Ex'] = [Notempty_Filter_Latest12.loc[Notempty_Filter_Latest12['BKH - Booking Ref'] == i,'VR - Rollover Date'] for i in Left_join['Job Reference']]

    for idx, rw in Left_join.iterrows():
        a1 = Left_join.loc[idx,'Modify date_Ex'].to_list()
        b1 = Left_join.loc[idx,'Start Date Time']
        if len(a1)>0:
            if a1[0] <b1:
                Left_join.loc[idx,'Rule 1'] = "True"
            else:
                Left_join.loc[idx,'Rule 1'] = "False"

    Left_join.loc[Left_join['Modify date']==0,'Rule 1'] = "False"
    Left_join.loc[(Left_join['Comment']=="Customer Rollover") & (Left_join['Rule 1']=="True") ,'Comment'] = "Modification date prior to Start Move"


    #ptn13.

    Left_join['Rule 2'] = ''
    dLatest = dfRO[['BKH - Booking Ref','VR - Rollover Date','VR - Fault']]
    dLatest =dLatest[dLatest['VR - Fault']=='CARRIER'] 
    dLatest['a_type_date'] = [min(dLatest[(dLatest['BKH - Booking Ref']==i)]['VR - Rollover Date']) for i in dLatest['BKH - Booking Ref']]

    Filter_Latest12 = dLatest[['BKH - Booking Ref','VR - Rollover Date','VR - Fault','a_type_date']]
    Notempty_Filter_Latest12 = Filter_Latest12[Filter_Latest12['a_type_date'].notnull()]
    Notempty_Filter_Latest12 = Notempty_Filter_Latest12[Notempty_Filter_Latest12['VR - Rollover Date'] == Notempty_Filter_Latest12['a_type_date']]
    Notempty_Filter_Latest12 = Notempty_Filter_Latest12.drop_duplicates()

    # Notempty_Filter_Latest12 = Notempty_Filter_Latest12[Notempty_Filter_Latest12['VR - Fault']=='CARRIER']
    Left_join['Modify date_Ex'] = [Notempty_Filter_Latest12.loc[Notempty_Filter_Latest12['BKH - Booking Ref'] == i,'VR - Rollover Date'] for i in Left_join['Job Reference']]

    for idx, rw in Left_join.iterrows():
        a1 = Left_join.loc[idx,'Modify date_Ex'].to_list()
        b1 = Left_join.loc[idx,'Start Date Time']
        if len(a1)>0:
            if a1[0] <b1:
                Left_join.loc[idx,'Rule 2'] = "True"
            else:
                Left_join.loc[idx,'Rule 2'] = "False"

    Left_join.loc[Left_join['Modify date']==0,'Rule 2'] = "False"
    Left_join.loc[((Left_join['Comment']=='Customer Rollover') + (Left_join['Comment']=='Modification date prior to Start Move')) & (Left_join['Rule 2']=="True"),'Comment'] = 'Prior to MOS Carrier Roll'

    # #ptn14.

    dLatest1 = dfRO[['BKH - Booking Ref','VR - Fault']]
    dLatest1 = dLatest1[dLatest1['VR - Fault'] == 'CUSTOMER']
    dLatest1.drop_duplicates(subset ="BKH - Booking Ref",
                        keep = False, inplace = True)

    Left_join['Matched_Bkg'] = [dLatest1.loc[dLatest1['BKH - Booking Ref'] == i,'BKH - Booking Ref'] for i in Left_join['Job Reference']]

    Left_join['Rule 3'] = ""

    for idx, rw in Left_join.iterrows():
        a1 = Left_join.loc[idx,'Matched_Bkg'].to_list()
        if len(a1)>0:
            pass
        else:
            Left_join.loc[idx,'Rule 3'] = "False"

    Left_join.loc[Left_join['Modify date']==0,'Rule 3'] = "False"
    Left_join.loc[(Left_join['Comment']=='Customer Rollover')  & (Left_join['Rule 3']=="False"),'Comment'] = 'Multiple Rollover'


    #ptn15.

    Left_join['Original booking'] = Left_join['Job Reference'].str[:10]

    #################### confirm #######################
    # dfRO = pd.read_excel("DDSM Roll over details.xlsx",sheet_name="Rollover")
    dfRO = pd.read_excel(filename,sheet_name="Rollover")
    #################### confirm #######################
    
    
    dLatest = dfRO[['BKH - Booking Ref','VR - Rollover Date','VR - Fault']]

    dLatest['a_type_date'] = [min(dLatest[(dLatest['BKH - Booking Ref']==i)]['VR - Rollover Date']) for i in dLatest['BKH - Booking Ref']]

    Filter_Latest13 = dLatest[['BKH - Booking Ref','VR - Rollover Date','VR - Fault','a_type_date']]
    # Filter_Latest13['BKH - Booking Ref'] = Filter_Latest13['BKH - Booking Ref'].str[:10]

    Notempty_Filter_Latest13 = Filter_Latest13[Filter_Latest13['a_type_date'].notnull()]
    empty_Filter = Filter_Latest13[Filter_Latest13['a_type_date'].isnull()]

    Notempty_Filter_Latest13 = Notempty_Filter_Latest13[(Notempty_Filter_Latest13['VR - Rollover Date'] == Notempty_Filter_Latest13['a_type_date']) & (Notempty_Filter_Latest13['a_type_date'].notnull())]
    Notempty_Filter_Latest13 = Notempty_Filter_Latest13.drop_duplicates()

    Left_join['Foult_3'] = [Notempty_Filter_Latest13.loc[Notempty_Filter_Latest13['BKH - Booking Ref'] == i,'VR - Fault'].values for i in Left_join['Original booking']]
    Left_join['Foult_13'] = [np.size(i) for i in Left_join['Foult_3']]
    Left_join.loc[df['Foult_13'] == 0 ,"Foult_3" ] = "0"

    lstvl = empty_Filter['BKH - Booking Ref'].to_list()
    Left_join.to_excel("pt15.xlsx")

    Left_join = pd.read_excel('pt15.xlsx')

    for id,rw in Left_join.iterrows():
        if Left_join.loc[id,"Foult_3"] == "['CUSTOMER']": 
            Left_join.loc[id,"Foult_3"] = 'CUSTOMER'
        elif Left_join.loc[id,"Foult_3"] == "['CARRIER']":
            Left_join.loc[id,"Foult_3"] = 'CARRIER'

    Left_join.loc[Left_join['Original booking'].isin(lstvl),"Foult_3"] = "xx"

    Left_join.loc[(Left_join['Comment']=="No Rollover – Split") & (Left_join['Foult_3']=="CARRIER"),'Comment'] = 'No Rollover – Split - Carrier'
    Left_join.loc[(Left_join['Comment']=="No Rollover – Split") & (Left_join['Foult_3']=="CUSTOMER"),'Comment'] = 'No Rollover – Split - Customer'
    Left_join.loc[(Left_join['Comment']=="No Rollover – Split") & (Left_join['Foult_3']=="xx"),'Comment'] = 'No Rollover – Split - Invoice'



    #RFI needs to be raise for booking status 1 
    Left_join.loc[(Left_join['status'] == 'Pending TBI') & (Left_join['Comment'].isna()) & (Left_join['Job Reference'].isna()) ,'Comment'] = "Correct booking needs to be link"
    Left_join.loc[(Left_join['status'] == 'Pending TBI') & (Left_join['Comment'].isna()) & (Left_join['Exp_Det Booking Status'] == 9) ,'Comment'] = "Correct booking needs to be link"


    Left_join.loc[(Left_join['status'] == 'Pending TBI') & (Left_join['Comment'].isna()) ,'Comment'] = "RFI needs to be raise for booking status 1"


########################################33
    Left_join.to_excel("OUTPUT_.xlsx",index=False)

    messagebox.showinfo("Process completed!!", "Thank you")




root = Tk()
root.title('Choose File')
root.geometry("700x200")
root.resizable(0,0)
root.config(background = "white")
label_file_explorer = Label(root,
                            text = "TBI Automation",
                            width = 100, height = 4,
                            fg = "blue")
button_explore = Button(root,
                        text = "Browse Files",
                        command = browseFiles)

button_Process = Button(root,
                        text = "Data Process_Export Detention",
                        command = OutProcess)


label_file_explorer.grid(column = 0, row = 1)
button_explore.grid(column = 0, row = 2)
button_Process.grid(column = 0, row = 8)



root.mainloop()

#%%
######################################################################
##Demerage part _____________________________________________________----
# ach***

import pandas as pd
import numpy as np
from tkinter import filedialog

def filePath(flName):
    filetype = (('Excel Files','*.xlsx'),('All files','*.*'))
    sourceFile = filedialog.askopenfile( mode='r', filetypes=filetype,title=flName)
    if sourceFile is not None:
        content = sourceFile.name
        return content

df = pd.read_excel(filePath('Rw Data'),sheet_name="Sheet1")
dfRO = pd.read_excel(filePath('Roll Over File'),sheet_name="Sheet1")
cutoff = pd.read_excel(filePath('Cutt off File'),sheet_name="Sheet1")
dfAlaska = pd.read_excel(filePath('Alska_Hawai'),sheet_name="ALASKA")
dfHW = pd.read_excel(filePath('Alska_Hawai'),sheet_name="HAWAII")

# # ptn2.
df = df[df['status']=='Pending TBI']

df['Comment'] = ""
df.loc[((~df['START_LOCATION'].str.startswith("CA") & ~df['START_LOCATION'].str.startswith("US"))),"Comment"] = "Overseas Pairing"
# # ptn3.
df.loc[((~df['Stop Location'].str.startswith("CA") & ~df['Stop Location'].str.startswith("US"))),"Comment"] = "Overseas Pairing"

## ptn4.
df.loc[((df['START_LOCATION'].str.startswith("CA")) & (df['Stop Location'].str.startswith("CA"))),"Comment"] = "Canada Pairing"

# # ptn5.
df.loc[((df['Start Status']=='IIT') & (df['Comment']=="")),"Comment"] = "Invalid Pairing -IIT to TPF"

# ptn6.
from datetime import timedelta

colmns = ['LFD','Cut-Off','Rule']
df[colmns] = ""
df['LFD'] = pd.to_datetime( df['Chargeable Date From']).dt.date - timedelta(days = 1)

df['Voyage Reference'].fillna(0,inplace=True)
cutoff.drop_duplicates(keep='first',inplace=True)

df.rename(columns={'Voyage Reference':'Voyage'},inplace=True)
cutoff['Uniq'] = cutoff['Voyage'].astype(str) +'|'+cutoff['Location'].astype(str)
df['Uniq'] = df['Voyage'].astype(str)+'|'+df['Stop Location'].astype(str)
cutoff['Uniq'] = cutoff['Uniq'].sort_values(axis=0,ascending=False, inplace=False)

df = pd.merge(df,cutoff[['Uniq','Cut-Off']],on='Uniq',how='left')
df.drop('Cut-Off_x',axis=1,inplace=True)
df['Cut-Off_y'].fillna('missing',inplace=True)

for id,rw in df.iterrows():
    if rw['Cut-Off_y'] !="missing":
        d2 = df.loc[id,'Cut-Off_y'].to_pydatetime().date()
        d1 = df.loc[id,'LFD']
        if d2<= d1:
            df.loc[id,'Rule'] = "True"
        else:
            df.loc[id,'Rule'] = "False"

df.loc[df['Rule']=='True','Comment'] = 'Cut-Off date Prior to LFD'


#ptn 7
colmns1 = ['Fault','Split','Original Booking','Door']
df[colmns1] = ""

df['Split'] = df['Job Reference'].str[-1:]
df['Door'] = df['Stop Location'].str[:5]
df['Original Booking'] = df['Job Reference'].str[:10]

for id, rw in df.iterrows():
    if df.loc[id,'Door'] ==df.loc[id,'Point Code']:
        df.loc[id,'Door'] = "True"
    else:
        df.loc[id,'Door'] = "False"



#ptn 8

dLatest = dfRO[['BKH - Booking Ref','VR - Rollover Date','VR - Fault']]
dLatest['a_type_date'] = [min(dLatest[(dLatest['BKH - Booking Ref']==i)]['VR - Rollover Date']) for i in dLatest['BKH - Booking Ref']]

Filter_Latest12 = dLatest[['BKH - Booking Ref','VR - Rollover Date','VR - Fault','a_type_date']]
Notempty_Filter_Latest12 = Filter_Latest12[Filter_Latest12['a_type_date'].notnull()]
Notempty_Filter_Latest12 = Notempty_Filter_Latest12[Notempty_Filter_Latest12['VR - Rollover Date'] == Notempty_Filter_Latest12['a_type_date']]
Notempty_Filter_Latest12 = Notempty_Filter_Latest12.drop_duplicates()

Notempty_Filter_Latest12.rename(columns={'BKH - Booking Ref':'Job Reference'},inplace=True)
df = pd.merge(df,Notempty_Filter_Latest12[['Job Reference','VR - Fault']],on='Job Reference',how='left')
df.drop('Fault',axis=1,inplace=True)

dfRO['VR - Fault'].fillna("Miss",inplace=True)

df.loc[(df['Comment'] =='') & (df['VR - Fault'] =='CUSTOMER'),'Comment'] = 'Customer Rollover'
df.loc[(df['Comment'] =='') & (df['VR - Fault'] =='CARRIER'),'Comment'] = 'Carrier Rollover'

missRollOver = dfRO[dfRO['VR - Fault'] =='Miss']
missRollOver.rename(columns={"BKH - Booking Ref":"Job Reference"},inplace=True)
df = df.merge(missRollOver[['Job Reference','VR - Fault']],on='Job Reference',how='left')

df.loc[((df['Split'].str.contains('[A-Za-z]',na=False)) & (df['VR - Fault_y'] =="Miss")& (df['Comment'] =="")),'Comment']='No Rollover - Split'
df.loc[(~df['Split'].str.contains('[A-Za-z]',na=False))  & (df['VR - Fault_y'] =="Miss")& (df['Comment'] ==""),"Comment"] = 'No Rollover'

df['Job Reference'] = df['Job Reference'].astype(str)



# pts 10
df.loc[(df['Job Reference'].str.len() <10) & (df['Comment'] ==''),'Comment'] = 'Need to be link Correct Booking'
# pts 11
df.loc[(df['Exp_Dem Booking Staus']==9) & (df['Comment'] ==''),'Comment'] = 'Need to be link Correct Booking'

# pts 12
df.loc[(df['Comment'] ==''),'Comment'] = 'Need to be RFI for Booking Status 1'




# ptn 13
clm = ['Modify date','Rule 1']
df[clm] = ""

dfRO['Revised_Date11'] = [max(dfRO[dfRO['BKH - Booking Ref']==i]['VR - Rollover Date']) for i in dfRO['BKH - Booking Ref']]
newestFilter = dfRO[dfRO['VR - Rollover Date']==dfRO['Revised_Date11']]
newestFilter.rename(columns={'BKH - Booking Ref':'Job Reference'},inplace=True)
df = pd.merge(df,newestFilter[['Job Reference','VR - Rollover Date']],how='left',on='Job Reference')
df['Rule 1'] = df['VR - Rollover Date']<df['Start Date Time']
df.loc[(df['Comment']=='Customer Rollover') & (df['Rule 1']==True)& (df['Origin']=='Port'),'Comment'] = 'Modification date prior to Start Move'
df.loc[(df['Comment']=='Customer Rollover') & (df['Rule 1']==True)& (df['Origin']=='Door')& (df['Door']=='True'),'Comment'] = 'Modification date prior to Start Move'


# pts 14

clm = ['Modify date for carrier','Rule 2']
df[clm] = ""
dfRO1= dfRO[dfRO['VR - Fault']=='CARRIER']
dfRO1['Revised_Date22'] = [min(dfRO1[dfRO1['BKH - Booking Ref']==i]['VR - Rollover Date']) for i in dfRO1['BKH - Booking Ref']]
newestFilter11 = dfRO1[dfRO1['VR - Rollover Date']==dfRO1['Revised_Date22']]
newestFilter11.rename(columns={'BKH - Booking Ref':'Job Reference'},inplace=True)

df = pd.merge(df,newestFilter11[['Job Reference','VR - Rollover Date']],how='left',on='Job Reference')

df['Rule 2'] = ((df['VR - Rollover Date_y']) < (df['Start Date Time']))

df.loc[((df['Comment']=='Customer Rollover') | (df['Comment']=='Modification date prior to Start Move')) & (df['Rule 2']==True)& (df['Origin']=='Port'),'Comment'] = 'Prior to XRX Carrier Roll'
df.loc[((df['Comment']=='Customer Rollover') | (df['Comment']=='Modification date prior to Start Move')) & (df['Rule 2']==True)& (df['Origin']=='Door')& (df['Door']=='True'),'Comment'] = 'Prior to XRX Carrier Roll'


# pts 15
df['Rule 3'] = ''
dfRO2 = dfRO[dfRO['VR - Fault'] =='CUSTOMER']
Without_dup = dfRO2.drop_duplicates(subset='BKH - Booking Ref',keep=False)
Without_dup.rename(columns={'BKH - Booking Ref':'Job Reference'},inplace=True)
Without_dup = Without_dup[Without_dup['VR - Fault'] == 'CUSTOMER']
df = pd.merge(df,Without_dup[['Job Reference','BKH - BL Number']],how='left',on='Job Reference')

for idx,i in df.iterrows():
    if type(i['BKH - BL Number']) == float:
        df.loc[idx,'BKH - BL Number'] = "Missing"
df.loc[(df['Comment']=='Customer Rollover') & (df['BKH - BL Number'] == 'Missing'),'Comment'] = 'Multiple Rollover'


# pts 16

clm = ['Split 1']
df[clm] = ""
dfRO['Revised_Date33'] = [min(dfRO[dfRO['BKH - Booking Ref']==i]['VR - Rollover Date']) for i in dfRO['BKH - Booking Ref']]
newestFilter12 = dfRO[dfRO['VR - Rollover Date']==dfRO['Revised_Date33']]
newestFilter12.rename(columns={'BKH - Booking Ref':'Original Booking'},inplace=True)

df = pd.merge(df,newestFilter12[['Original Booking','VR - Fault']],how='left',on='Original Booking')

df.loc[(df['Comment'] =='No Rollover - Split')  & (df['VR - Fault'] =='CARRIER'),'Comment'] = 'No Rollover – Split - Carrier'
df.loc[(df['Comment'] =='No Rollover - Split')  & (df['VR - Fault'] =='CUSTOMER'),'Comment'] = 'No Rollover – Split - Customer'


missing = dfRO[dfRO['VR - Fault']=='Miss'][['BKH - Booking Ref','VR - Fault']]
missing.rename(columns={'VR - Fault':'Null_Status','BKH - Booking Ref':'Original Booking'},inplace=True)

df = pd.merge(df,missing[['Original Booking','Null_Status']],how='left',on='Original Booking')


df.loc[(df['Comment'] =='No Rollover - Split')& (df['VR - Fault_y']=='Miss'),'Comment'] = 'No Rollover – Split - Invoice'


# pts 17
df.loc[((df['Comment']=='No Rollover')| (df['Comment']=='No Rollover – Split - Invoice') | (df['Comment']=='No Rollover – Split - Customer')) & (df['Origin'] =='Inland Point / Ramp'),'Comment'] = 'Rail Shipment'
df.loc[((df['Comment']=='No Rollover')| (df['Comment']=='No Rollover – Split - Invoice') | (df['Comment']=='No Rollover – Split - Customer')) & (df['Origin'] =='Door') &  (df['Door'] =='False'),'Comment'] = 'Rail Shipment'

#%%

# pts 18
# dfAl = dfAlaska['CMA POOL']
dfAlaska.rename(columns={'CMA POOL':'START_LOCATION'},inplace=True)
df = pd.merge(df,dfAlaska[['START_LOCATION','State']],on='START_LOCATION',how='left')

dfAlaska.rename(columns={'START_LOCATION':'Stop Location'},inplace=True)
df = pd.merge(df,dfAlaska[['Stop Location','CMA POOL NAME']],on='Stop Location',how='left')

df['CMA POOL NAME'].astype(str)
df['State'].astype(str)

df.loc[df['State'].notnull(),'Comment'] ='Alaska Location'
df.loc[df['CMA POOL NAME'].notnull(),'Comment'] ='Alaska Location'

#%%
# pts 19
dfHW.rename(columns={'POOL LOCATION':'START_LOCATION'},inplace=True)
df = pd.merge(df,dfHW[['START_LOCATION','FULL NAME']],on='START_LOCATION',how='left')

dfHW.rename(columns={'START_LOCATION':'Stop Location'},inplace=True)
df = pd.merge(df,dfHW[['Stop Location','ISLAND']],on='Stop Location',how='left')

df.loc[(df['FULL NAME'].notnull()) & (df['Comment']=='Customer Rollover')& (df['Comment']=='Modification date prior to Start Move')& (df['Comment']=='Multiple Rollover')& (df['Comment']=='No Rollover')& (df['Comment']=='No Rollover – Split')& (df['Comment']=='No Rollover – Split – Customer')& (df['Comment']=='No Rollover – Split – Invoice'),'Comment'] ='Hawaii Location'
df.loc[(df['ISLAND'].notnull())& (df['Comment']=='Customer Rollover')& (df['Comment']=='Modification date prior to Start Move')& (df['Comment']=='Multiple Rollover')& (df['Comment']=='No Rollover')& (df['Comment']=='No Rollover – Split')& (df['Comment']=='No Rollover – Split – Customer')& (df['Comment']=='No Rollover – Split – Invoice'),'Comment'] ='Hawaii Location'

# testing
#%%


# #
# #%%

# import pandas as pd

# df = pd.read_excel("OUTPUT_.xlsx")
# print('Done')
# #%%


# #%%
# df.loc[(df['status'] == 'Pending TBI') & (~df['XOF'].isna()) & (df['Exp_Det Booking Status'] == '1'),'Comments'] = "RFI needs to be raise for booking status 1"
