# #%%
# import pandas as pd
# from datetime import datetime,date
# import win32com.client as win32
# from dateutil import parser
# from tkinter import *
# import tkinter as tk
# import tkinter.font as font
# from tkinter import messagebox
# import win32timezone

# # 22
# gui = Tk(className='Outlook Extraction')
# gui.geometry("900x200")

# outlook = win32.Dispatch('Outlook.Application')
# namespace = outlook.GetNamespace("MAPI")
# list_Subj = []
# list_Rec = []
# list_From = []
# list_to= []

# def autoPick():
    
#     pFolder = namespace.PickFolder()
#     tFolder = pFolder.Items


#     for i in tFolder:
#         try:
#             subj = i.Subject
#             dTime = i.SentOn.strftime('%d/%m/%Y %H:%m')
#             FrmName = i.SenderName
#             sTo = i.To


#             list_Subj.append(subj)
#             list_Rec.append(dTime)
#             list_From.append(FrmName)
#             list_to.append(sTo)
#             df = pd.DataFrame({"Subject":list_Subj, "RecievedTime" : list_Rec,"From":list_From,"To":list_to})            
#         except AttributeError:
#             pass

#     df.to_excel('Mail_list.xlsx',index=False)

#     messagebox.showinfo("Done",'Extraction Completed ..')


# def extrM():
#     # outlook = win32.Dispatch('Outlook.Application')

#     # namespace = outlook.GetNamespace("MAPI")
#     accnt = namespace.Folders(entry.get()).Folders('Inbox')
#     tFolder = accnt.Folders(entry2.get()).Items

#     # list_Subj = []
#     # list_Rec = []
#     # list_From = []
#     # list_to= []    
#     for i in tFolder:
#         try:
#             subj = i.Subject
#             dTime = i.SentOn.strftime('%d/%m/%Y %H:%m')
#             FrmName = i.SenderName
#             sTo = i.To


#             list_Subj.append(subj)
#             list_Rec.append(dTime)
#             list_From.append(FrmName)
#             list_to.append(sTo)
#             df = pd.DataFrame({"Subject":list_Subj, "RecievedTime" : list_Rec,"From":list_From,"To":list_to})            
#         except AttributeError:
#             pass

#     df.to_excel('Mail_list.xlsx',index=False)

#     messagebox.showinfo("Done",'Extraction Completed ..')

# myFont = font.Font(family='Helvetica', size=18, weight='bold')

# lbl = tk.Label(gui, text='Email ID')
# lbl['font'] = myFont
# lbl.place(relx=0.0, rely=0.1, anchor='w')


# entry = Entry(gui, width= 42)
# entry.place(relx= 0.3, rely= 0.1, anchor= CENTER)



# lbl = tk.Label(gui, text='Folder')
# lbl['font'] = myFont
# lbl.place(relx=0.5, rely=0.1, anchor='w')

# entry2 = Entry(gui, width= 42)
# entry2.place(relx= 0.8, rely= 0.1, anchor= CENTER)

# button = Button(gui, text='GetMail', bg='#0052cc', fg='#ffffff',height= 2, width=10,command=extrM)
# button['font'] = myFont
# button.place(relx=0.4, rely=0.5, anchor=CENTER)

# button1 = Button(gui, text='Auto_Pick_Folder', bg='#0052cc', fg='#ffffff',height= 2, width=20,command=autoPick)
# button1['font'] = myFont
# button1.place(relx=0.7, rely=0.5, anchor=CENTER)


# gui.mainloop()
# cDate = date.today()

#%%
# cDateMail = [msg.EntryID for msg in inbox if msg.SentOn.strftime('%d-%m-%Y')==cDate.strftime('%d-%m-%Y')]

# from email.header import Header
# from email.utils import formataddr

# for i in cDateMail:
#     # namespace.GetItemFromID(i).display()
#     mail = namespace.GetItemFromID(i)

#%%
# =================

# import pandas as pd
# from tkinter import filedialog
# from tkinter import *
# import tkinter.font as font

# gui = Tk(className='Report')
# gui.geometry("500x200")


# def browse(ttl):
#     fl = filedialog.askopenfile(title=ttl)
#     return fl.name


# def proce():
#     # df = pd.read_excel('ssc.dgsupport.eur@cma-cgm.com.xlsx')
#     df = pd.read_excel(browse('Extraction_File'))
#     df1 = df[['Subject','DisplayTo','DateTimeReceived']]
#     CurYear = df1[df1['DateTimeReceived'].astype('str').str.contains('2022')]

#     conDF = pd.read_excel(browse('Consolidate'))
#     # conDF1 = conDF[conDF['DCO']=='Europe']
#     # conDF1 = conDF[conDF['DCO']=='Asia']
#     conDF1 = conDF
#     # drop duplicate
#     conDF1 = conDF1.drop_duplicates(subset='Booking Number',keep=False)

#     for id, rw in conDF1.iterrows():
#         if pd.notnull(conDF1.loc[id,'Booking Number']):
#             # bkg_nu = "26045175"
#             try:
#                 bkg_nu = str(conDF1.loc[id,'Booking Number'])
#                 print(bkg_nu)
#                 fndBkg = CurYear[CurYear['Subject'].astype('str').str.contains(bkg_nu)]
#                 if len(fndBkg)>0:                
#                     startMail = fndBkg[fndBkg['DateTimeReceived'] == min(fndBkg['DateTimeReceived'])]
#                     conDF1.loc[id,'Consolidation_Data']= startMail['DateTimeReceived'].to_string(index=False)
#             except Exception:
#                 pass
#     # conDF1.to_excel('Asia.xlsx',index=False)
#     conDF1.to_excel('Out.xlsx',index=False)

# myFont = font.Font(family='Helvetica', size=20, weight='bold')

# button1 = Button(gui, text='Auto_Pick_Folder', bg='#0052cc', fg='#ffffff',height= 2, width=15,command=proce)
# button1['font'] = myFont
# button1.place(relx=0.5, rely=0.5, anchor=CENTER)


# gui.mainloop()
#%%



# ff  = [min(CurYear[CurYear['Subject'].astype('str').str.contains('571200042043')]['DateTimeReceived'])]
# fstDate = ff[0].to_pydatetime().date().strftime('%Y-%m-%d') 

#%%


#%%

# # TAT CALCULATION
import pandas as pd
import dateutil.parser

# import pandas as pd
from tkinter import filedialog
from tkinter import *
import tkinter.font as font

gui = Tk(className='Report')
gui.geometry("500x200")


def browse2(ttl):
    fl = filedialog.askopenfile(title=ttl)
    return fl.name


def proc():
    df = pd.read_excel(browse2('OutPut_File'))

    for id,rw in df.iterrows():
        try:
            strTime = df.loc[id,'Start']
            endTime = df.loc[id,'End']
            print(strTime)
            print(endTime)
            if pd.notnull(strTime):
                stDate = pd.to_datetime(strTime)
                endDate = pd.to_datetime(endTime)

                dTimeDiff = endDate-stDate
                dayC = dTimeDiff.days
                hrsT = dTimeDiff.seconds/3600
            df.loc[id,'Days'] = dayC
            df.loc[id,'Hrs_Minute'] = round(hrsT,2)
        except Exception:
            pass
    df.to_excel('out.xlsx',index=False)

myFont = font.Font(family='Helvetica', size=20, weight='bold')

button1 = Button(gui, text='TAT_Calc', bg='#0052cc', fg='#ffffff',height= 2, width=15,command=proc)
button1['font'] = myFont
button1.place(relx=0.5, rely=0.5, anchor=CENTER)


gui.mainloop()

#%%



# import pandas as pd
# import dateutil.parser

# # import pandas as pd
# from tkinter import filedialog
# from tkinter import *
# import tkinter.font as font

# gui = Tk(className='Report')
# gui.geometry("500x200")


# def browse2(ttl):
#     fl = filedialog.askopenfile(title=ttl)
#     return fl.name

# def proc():    
#     df_C = pd.read_excel(browse2('OutPut_File'))


#     for id,rw in df_C.iterrows():
#         try:
#             if pd.isnull(df_C.loc[id,'Consolidation_Data']):
#                 cDate = df_C.loc[id,'Email Receive date'].strftime("%Y-%m-%d")   
#                 end_T= df_C.loc[id,'Input start time']
#                 Combine_DateTime =  cDate +" "+ end_T
#                 NearVal = dateutil.parser.parse(Combine_DateTime)
#                 df_C.loc[id,'Consolidation_Data'] = NearVal.strftime("%Y-%m-%d %H:%M:%S")
#                 print('updating_Blank')
#         except Exception:
#             pass
#     df_C.to_excel("out.xlsx",index=False)

# myFont = font.Font(family='Helvetica', size=20, weight='bold')

# button1 = Button(gui, text='Cleaning_Empty', bg='#0052cc', fg='#ffffff',height= 2, width=15,command=proc)
# button1['font'] = myFont
# button1.place(relx=0.5, rely=0.5, anchor=CENTER)

# gui.mainloop()
#%%

# import pandas as pd
# from datetime import datetime,date

# from tkinter import filedialog
# from tkinter import *
# import tkinter.font as font

# gui = Tk(className='OUT_Report')
# gui.geometry("500x200")


# def browse2(ttl):
#     fl = filedialog.askopenfile(title=ttl)
#     return fl.name

# def proc():
#     # df = pd.read_excel('ANL.xlsx')
#     df = pd.read_excel(browse2('Extracted_File'))    
#     df = df[df['DateTimeReceived'].astype('str').str.contains('2022')]

#     # df_A = pd.read_excel('ANL_A.xlsx')
#     df_A = pd.read_excel(browse2('Out_File'))    
#     df_A = df_A[~df_A['Activity'].astype('str').str.contains('LARA EDI')]


#     for id, rw in df_A.iterrows():
#         try:
#             bkg = df_A.loc[id,'Booking Number']    
#             if type(bkg) == int:
#                 bkg = str(bkg)
#             else:
#                 pass
#             bkg_ = bkg.strip()
#             print(bkg_)
#             # bkg_ = 'akhileshcha-uhan'
#             nLst = []
            
#             lst = ['|','-','#','*','?']
#             for itm in lst:
#                 if itm in bkg_:                
#                     bkg_N = bkg_.split(itm)[0]
#                     nLst.append(bkg_N)
#                     break   

#             if len(nLst)>0:
#                 bkg_ = nLst[0]         
#             else:
#                 bkg_ = bkg.strip()

            
#             foundVal = df[df['Subject'].astype('str').str.contains(bkg_)]
#             sorted_df = foundVal.sort_values(by=['DateTimeReceived'], ascending=True)

#             lst = sorted_df['DateTimeReceived'].astype('str').to_list()
#             sentDate = sorted_df.loc[sorted_df['Folder Path'].astype('str').str.contains('Sent Items'),'DateTimeReceived'].to_string(index=False)
#             if "\n" in sentDate:
#                 sentDate = sentDate.split("\n")[0]
#             else:
#                 pass
                
#             position = lst.index(sentDate)
#             if position >=1:
#                 indx = (position-1)
#                 recived_Date = lst[indx]
#             df_A.loc[id,"Start"] =  recived_Date
            
#             df_A.loc[id,"End"] =  sentDate
#         except Exception:
#             pass
#     df_A.to_excel('out.xlsx',index=False)

# myFont = font.Font(family='Helvetica', size=20, weight='bold')

# button1 = Button(gui, text='Outlook_Data', bg='#0052cc', fg='#ffffff',height= 2, width=15,command=proc)
# button1['font'] = myFont
# button1.place(relx=0.5, rely=0.5, anchor=CENTER)


# gui.mainloop()

# #%%

# import pandas as pd

# df = pd.read_excel("ANL_A.xlsx")
# #%%
# for id,rw in df.iterrows():
#     try:
#         strTime = df.loc[id,'Start']
#         endTime = df.loc[id,'End']
#         print(strTime)
#         print(endTime)
#         if pd.notnull(strTime):
#             stDate = pd.to_datetime(strTime)
#             endDate = pd.to_datetime(endTime)

#             dTimeDiff = endDate-stDate
#             dayC = dTimeDiff.days
#             hrsT = dTimeDiff.seconds/3600
#         df.loc[id,'Days'] = dayC
#         df.loc[id,'Hrs_Minute'] = round(hrsT,2)
#     except Exception:
#         pass
# df.to_excel('out.xlsx',index=False)
    