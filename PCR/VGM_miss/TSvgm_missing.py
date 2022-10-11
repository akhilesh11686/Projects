#%%
import pandas as pd
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



df_6 = pd.read_excel('LCR06.xlsx',skiprows=17)
df_11 = pd.read_excel('LCR11.xlsx',skiprows=17)

lUnt = df_11[df_11['PTS Code']==df_11['Final POD']].index
df_11.drop(lUnt,inplace=True)

# exclude 1|9
excl_1_9 =df_11[~df_11['Booking Status'].astype('str').str.contains('1|9')]

#exclude empty flage : Y
NonEmpty = excl_1_9[excl_1_9['Empty Flag']=='N']

# empty vgm Verify Gross Mass
missVGM = NonEmpty[NonEmpty['Verify Gross Mass'].isnull()]

#%%
import pandas as pd



df_6 = pd.read_excel('LCR06.xlsx',skiprows=17)

lUnt = df_6[df_6['Port']==df_6['First POL']].index
df_6.drop(lUnt,inplace=True)

lUnt_1 = df_6[df_6['Port']==df_6['Final POD']].index
df_6.drop(lUnt_1,inplace=True)

#exclude empty flage : Y
NonEmpty = df_6[df_6['Empty']=='N']

# empty vgm Verify Gross Mass
missVGM = NonEmpty[NonEmpty['Verified Gross Mass'].isnull()]
