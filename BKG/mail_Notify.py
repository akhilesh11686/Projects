
#%%
import pandas as pd
from openpyxl import load_workbook
from tkinter import filedialog

def browseFile():
    path_f = filedialog.askopenfile()
    return path_f


# pth = browseFile()

xl = pd.ExcelFile('BKG_DCD.xlsm',engine='openpyxl')
xl_emlTb = xl.parse(sheet_name='Email_Tbl')
xl_ChngDepo = xl.parse(sheet_name='Change of Depot')
#%%
