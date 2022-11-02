#%%

import pandas as pd
from tkinter import filedialog
from datetime import date

def browse(ttl):
    pth = filedialog.askopenfile(title=ttl)
    return pth


fVsl = browse('vessel Schedule')
Excp = browse('Exception_File')

df = pd.read_excel(fVsl.name,skiprows=2)
dfExc = pd.read_excel(Excp.name)

# Blank row
dfNBlnk= df.loc[~df['Port Code'].isnull(),['Port Code','Port Name','Voyage Ref']]

# PortCode with PortName
portCode = dfNBlnk[~dfNBlnk['Port Name'].isnull()]

cDate = date.today().strftime('%d-%m-%Y')

# df[df['Cutoff'].astype(str).str.contains([cDate])]

print('Found')
#%%
df = df[(~df['Voyage Ref'].isnull()) & (~df['Voyage Ref'].astype(str).str.endswith('PL'))]
