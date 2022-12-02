#%%
from tkinter import *
import tkinter.font as font
from tkinter import filedialog
import pandas as pd
import os
from datetime import date,timedelta
import re
from tkinter import messagebox

gui = Tk(className='DPS report')
gui.geometry("600x200")



def browseF(Dfl):
    fl = filedialog.askopenfile(title= Dfl)
    return fl.name

def process():
    df11 = browseF('DPS file')
    df = pd.read_excel(df11)
    rowCnt,ClmnCnt = df.shape
    for m in range(0,ClmnCnt):
        for n in range(0,rowCnt):
            if df.iloc[n,m]=='ETA':
                clmETA=m
                break
            else:
                pass


    for p in range(0,ClmnCnt):
        for q in range(0,rowCnt):
            if df.iloc[q,p]=='Port':
                clmVsl=q
                break
            elif df.iloc[q,p]=='Vessel':
                clmVsl=q
                break  
            else:
                pass          

    dInclu = int(E1.get())
    dateList1 = []

    if dInclu == 1:
        today = (date.today()+timedelta(days=1))
        d1 = today.strftime("%d/%m")    
    elif dInclu > 1:
        today = (date.today()+timedelta(days=dInclu))
        for k in range(1,dInclu+1):
            mm = date.today()+timedelta(days=k)
            d1 = mm.strftime("%d/%m")
            dateList1.append(d1)





    CVslList = []
    for j in range(0,rowCnt):
            # get match ETA
            if df.iloc[j,clmETA]=='ETA':
                vsl = df.iloc[j,0]
                for i in range((clmETA+1),ClmnCnt):                
                    cVal = str(df.iloc[j,i])

                    if len(dateList1)>0:
                        for x in dateList1:
                            if x in cVal:
                                Ports = df.iloc[clmVsl,i]
                                dTimeVal = df.iloc[j,i]
                                # unkVal = vsl +"|"+ Ports + "|" + dTimeVal +"|"+ 'Rows :' + str(j)+"|"+ 'Columns :' + str(i)
                                unkVal = vsl +"|"+ Ports + "|" + dTimeVal 
                                CVslList.append(unkVal)
                            else:
                                pass

                    elif len(dateList1)==0:
                        if d1 in cVal:
                            Ports = df.iloc[clmVsl,i]
                            dTimeVal = df.iloc[j,i]
                            # unkVal = vsl +"|"+ Ports + "|" + dTimeVal +"|"+ 'Rows :' + str(j)+"|"+ 'Columns :' + str(i)
                            unkVal = vsl +"|"+ Ports + "|" + dTimeVal 
                            CVslList.append(unkVal)
                        else:
                            pass
            else:
                pass
        

    fdf = pd.DataFrame({"Details":CVslList })
    splVal = fdf['Details'].str.split("|",n=1, expand= True)
    fdf['Vsl_portDetails1'] = splVal[0]

    fdf['Vsl_portDetails2'] = splVal[1].str.split("|",n=1, expand= True)[0]
    fdf['ETA_Time'] = splVal[1].str.split("|",n=1, expand= True)[1]

    fdf.to_excel('Out.xlsx',index=False)
    messagebox.showinfo('Done','Completed the date extraction')


L1 = Label(gui, text="Enter the number for addition in date")
L1.place( relx=0.5, rely=0.5, anchor='center')
L1.grid(columnspan=9,row=2,column=2)

E1 = Entry(gui, bd =5)
E1.place(relx=0.5, rely=0.5, anchor='center')
E1.grid(columnspan=9,row=2,column=20)

myFont = font.Font(family='Helvetica', size=90, weight='bold')
button = Button(gui, text='Get DPS', bg='#0052cc', fg='#ffffff',command=process,height=3, width=30)

button.grid(columnspan=9,row=5,column=2)
button.place(relx=0.5, rely=0.5, anchor='center')
gui.mainloop()




