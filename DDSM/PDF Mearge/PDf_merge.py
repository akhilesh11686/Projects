
#%%

from PyPDF2 import PdfFileMerger
import os
import pandas as pd
import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox

root = tk.Tk(className=' PDF Combining..')
root.geometry("700x200")
root.maxsize("700","200")

myfont = tkFont.Font(family='Helvetica',size=12)

def getMerge():
    lst = []

    if not os.path.exists('Output'):
        os.makedirs('Output')
        
    df  = pd.read_excel('BL_Invoice.xlsx',sheet_name='Combining_PDF')
    for i,rw in df.iterrows():
        try:
            lst.append(rw[0])
            lst.append(rw[1])
            merger = PdfFileMerger()
            pdf_files = [rw[0] +'.pdf', rw[1] +'.pdf']
            for pdf_file in pdf_files:
                #Append PDF files
                merger.append(pdf_file)

            # os.chdir('./Output')    
            merger.write('Output//'+rw[2]+'.pdf')
            # os.chdir('../')
            merger.close()
            lst.clear()
        except Exception:
            continue
    messagebox.showinfo("PDF Merge successfully",'Thank you !!' )    

buttonExample1 = tk.Button(root,
                           text="Click here to Merge PDF",
                           width=35,
                           height=4,command=getMerge,bg='Green',fg='White')
buttonExample1['font'] = myfont
buttonExample1.grid(padx=180,pady=50)

root.mainloop()

#%%

# from ast import Continue
# from csv import excel
# import win32com.client
# import os
# from tkinter import *
# import tkinter.font as font
# import tkinter.messagebox

# root = Tk()
# root.geometry("550x150")
# root.title('PDF extraction..')
# root.maxsize(False,False)

# # Acnt = 'ssc.achauhan@cma-cgm.com'
# # Fldr = 'Inbox'
# # Sub_Fldr = '3.Done'


# Acnt = 'ssc.importvipaudit@cma-cgm.com'
# Fldr = 'Inbox'
# Sub_Fldr = 'Invoices'


# def outlook_download():
    
#     global final_file_Downloaded
    
#     outlook = win32com.client.dynamic.Dispatch("Outlook.Application").GetNamespace("MAPI")
#     inbox = outlook.Folders[Acnt].Folders[Fldr].Folders[Sub_Fldr]
#     item_count =inbox.Items.Count
#     messages=inbox.Items
#     messages.Sort("[ReceivedTime]", True)
#     outputDir = os.getcwd()
#     # outputDir.mkdir(parents=True,exist_ok=True)


#     ext = '*.pdf'
#     for message in messages:
#         if message.Subject.find('B/L Draft:') >-1:
#             attachments = message.Attachments
#             m_attach = len([i for i in attachments])
#             for i in range(1, (m_attach+1)):
#                 attachment = attachments.Item(i)
#                 if (attachment.FileName).endswith('.pdf'):
#                         oldName = attachment.FileName
#                         lst = message.Subject.split(":")
#                         bln = lst[len(lst)-1]
#                         bln = bln.lstrip()
#                         newName = bln + ".pdf"
#                         try:
#                             attachment.SaveAsFile(os.path.join(outputDir,attachment.FileName))
#                             f1 = os.path.join(outputDir,attachment.FileName)
#                             f2 = os.path.join(outputDir,newName)
#                             os.rename(f1,f2)                        
#                             print("PDF file has been downloaded",attachment.FileName)
#                         except:
#                             continue

                        
#         else:

#             attachments = message.Attachments
#             m_attach = len([i for i in attachments])
#             for i in range(1, (m_attach+1)):
#                 attachment = attachments.Item(i)
#                 if (attachment.FileName).endswith('.pdf'):
#                         try:
#                             attachment.SaveAsFile(os.path.join(outputDir,attachment.FileName))
#                             print("PDF file has been downloaded",attachment.FileName)
#                         except:
#                             continue
                        
                        


# myFont = font.Font(family='Helvetica', size=20, weight='bold')
# Btn = Button(root,text='Extract PDF',bg='Blue',font='yellow',command=outlook_download)
# Btn['font'] = myFont
# Btn.grid(padx=200,pady=40)
# # outlook_download()
# root.mainloop()


# #%%
# from PyPDF2 import PdfFileMerger
# import os
# import pandas as pd
# import tkinter as tk
# import tkinter.font as tkFont
# from tkinter import messagebox

# lst = []

# if not os.path.exists('Output'):
#     os.makedirs('Output')



# df  = pd.read_excel('BL_Invoice.xlsx',sheet_name='Combining_PDF')
# for i,rw in df.iterrows():
#     lst.append(rw[0])
#     lst.append(rw[1])
#     merger = PdfFileMerger()
#     pdf_files = [rw[0] +'.pdf', rw[1] +'.pdf']
#     for pdf_file in pdf_files:
#         #Append PDF files
#         merger.append(pdf_file)

#     # os.chdir('./Output')    
#     merger.write('Output//'+rw[2]+'.pdf')
#     # os.chdir('../')
#     merger.close()
#     lst.clear()
# # pdf = pikepdf.Pdf.new()

# # for file in lst:
# #     src = pikepdf.Pdf.open(file + '.pdf')
# #     pdf.pages.extend(src.pages)    
    
# # os.chdir('./Output')
# # fName = rw[2]+'.pdf'  
# # pdf.save( fName)
# # os.chdir('../')
# messagebox.showinfo("PDF Merge successfully",'Thank you !!' )