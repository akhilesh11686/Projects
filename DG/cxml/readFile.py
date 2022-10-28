#%%
from ast import Break
from multiprocessing.resource_sharer import stop
import tkinter
import pandas as pd
from tkinter import *
from tkinter import messagebox
import tkinter.filedialog
import os



root = Tk()
root.geometry("300x100")
root.title('Cxml_to_xls')
def bro(fl):
    fl = tkinter.filedialog.askopenfile(title=fl)
    return fl


def proc():    
    fls = bro("Cxml File")
    xl = pd.read_xml(fls.name)

    hrpFile = bro("Harp file")
    Hrp = pd.read_excel(hrpFile.name)

    cnt = bro("Container map file")
    cntDf = pd.read_excel(cnt.name)


    xl['Value'] = xl['Value'].replace('Y','')


    aa = xl['EID'].drop_duplicates()
    df = pd.DataFrame()
    for i in aa:
        cls = xl.loc[xl['EID']==i,'Class'].dropna().to_string(index=False)
        un = xl.loc[xl['EID']==i,'UNNo'].dropna().to_string(index=False)
        pgrp = xl.loc[xl['EID']==i,'PackingGroup'].dropna().to_string(index=False)
        cntr = xl.loc[xl['EID']==i,'IdNumber'].dropna().to_string(index=False)
        Ctype = xl.loc[xl['EID']==i,'ISOType'].dropna().to_string(index=False)
        pols = xl.loc[xl['EID']==i,'POL'].dropna().to_string(index=False)
        pods = xl.loc[xl['EID']==i,'POD'].dropna().to_string(index=False)
        bkg = xl.loc[xl['EID']==i,'Value'].dropna().to_string(index=False)
        bkg_N = bkg
        df = df.append({'Booking': bkg_N, 'Container': cntr, 'POL': pols,'POD': pods,'Class' : cls, "UN No" : un, 'Group' : pgrp, 'Type': Ctype},ignore_index=True)    

    # split of multiple UN number
    df.replace('Series([], )',"-",inplace=True)    

    df['UN No'] = df['UN No'].str.replace('\n',",",)
    df['Class'] = df['Class'].str.replace('\n',",",)
    df['Group'] = df['Group'].str.replace('\n',",",)

    df = (df.apply(lambda x: x.str.split(',').explode()).reset_index())


    df.to_excel('output.xlsx',index=False)

    Hrp.columns = Hrp.columns.str.strip().str.lower().str.replace('\n', '_').str.replace('(', '').str.replace(')', '')

    for i,id in df.iterrows():
        if df.loc[i,'UN No']!="-":        
            # vl = Hrp[(Hrp['container_ no.']==id['Container']) & (Hrp['discharge _port']==id['POD']) & (Hrp['unno'].astype(str)==str(id['UN No'])) ]
            lst = cntDf.loc[cntDf['Container_size']==id['Type'],'Map'].to_list()[0]
            vl = Hrp[(Hrp['container_ no.']==id['Container']) & (Hrp['discharge _port']==id['POD']) & (Hrp['unno'].astype(str)==str(id['UN No'])) & (Hrp['container type'].str.contains(lst,regex=True,na=True)) ]
            # vl1 = (vl['container type'].str.contains(lst,regex=True,na=True))
            if len(vl)>0:
                if vl['operator'].to_string(index=False).strip()!='CMA':
                    df.loc[i,'Status'] = 'OK'
                else:
                    df.loc[i,'Status'] = 'Out of scope'
            else:
                df.loc[i,'Status'] = 'NOT OK'

    df.to_excel('final_status1.xlsx')



    messagebox.showinfo('Done!','completed process')


button = Button(root, text='Cxml_to_xls', bg='#0052cc', fg='#ffffff',command=proc)
button.pack()

root.mainloop()
#%%
import pandas as pd
from tkinter import filedialog
df = pd.read_excel('output.xlsx')
cntDf = pd.read_excel('Container_map.xlsx')
Hrp = pd.read_excel('LOTUS A 0PE4UE1MA BEANR Revised.xls')

Hrp.columns = Hrp.columns.str.strip().str.lower().str.replace('\n', '_').str.replace('(', '').str.replace(')', '')
#%%
# note , container type T!
for i,id in df.iterrows():
    if df.loc[i,'UN No']!="-":        
        # vl = Hrp[(Hrp['container_ no.']==id['Container']) & (Hrp['discharge _port']==id['POD']) & (Hrp['unno'].astype(str)==str(id['UN No'])) ]
        lst = cntDf.loc[cntDf['Container_size']==id['Type'],'Map'].to_list()[0]
        vl = Hrp[(Hrp['container_ no.']==id['Container']) & (Hrp['discharge _port']==id['POD']) & (Hrp['unno'].astype(str)==str(id['UN No'])) & (Hrp['container type'].str.contains(lst,regex=True,na=True)) ]
        # vl1 = (vl['container type'].str.contains(lst,regex=True,na=True))
        if len(vl)>0:
            if vl['operator'].to_string(index=False).strip()!='CMA':
                df.loc[i,'Status'] = 'OK'
            else:
                df.loc[i,'Status'] = 'Out of scope'
        else:
            df.loc[i,'Status'] = 'NOT OK'

df.to_excel('final_status1.xlsx')


