#%%
import tkinter
import pandas as pd
from tkinter import *
from tkinter import messagebox
import tkinter.filedialog


root = Tk()
root.geometry("300x100")
root.title('Cxml_to_xls')
def bro():
    fl = tkinter.filedialog.askopenfile()
    return fl


def proc():
    fls = bro()
    # xl = pd.read_xml('0PE4VW1MA-INNSA AGW LIST.cxml')
    xl = pd.read_xml(fls.name)
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

    df.to_excel('output.xlsx',index=False)
    messagebox.showinfo('Done!','completed process')

button = Button(root, text='Cxml_to_xls', bg='#0052cc', fg='#ffffff',command=proc)
button.pack()

root.mainloop()