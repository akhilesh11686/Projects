
#%%
# ttod
import tkinter
from tkinter import messagebox
import pandas as pd
from openpyxl import workbook
from  openpyxl import load_workbook

from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile

import openpyxl
import datetime
from datetime import date
import win32com.client as win32
import os

from tkinter.filedialog import askdirectory
from PIL import Image,ImageTk



root = Tk()
root.geometry('650x200')
root.resizable(False,False)
root.title('Mails_Distr..')


def open_file():
    global RwFilePath,mis_tbl,vyg,file
    
    file = askopenfile(mode='r',filetypes=[('Excel Files','*.XLSX')])
    RwFilePath = file.name
    pthLbl.configure(text = RwFilePath, font= ('Helvetica 10 italic'))

def png_Jpg():

    dirc = askdirectory(title="Select PNG folder")
    for fl in os.listdir(dirc):
        try:
            if fl.find(".png") or fl.find(".PNG"):
                img_png = Image.open(fl)
                lst = fl.split('.')
                if len(lst)>0:
                    img_png.save(lst[0]+ '.jpg')

        except:
            pass

    messagebox.showinfo("Thank you!!","Completed..")    


def Mail_dist():

    df = pd.read_excel(file.name,sheet_name='Mailing_list',dtype=str)
    # df_Content = pd.read_excel(file.name,sheet_name='Mail_body')

    outlook = win32.Dispatch('outlook.application')
    oacctouse = None
    for index,row in df.iterrows():

        frm = df.loc[index,'From Email address']
        eml = df.loc[index,'Emp_Email_address']
        subj = df.loc[index,'Subject']
        emp_ID = df.loc[index,'Emp_ID']

        for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == frm:
                oacctouse = oacc
                break

        outlookNm = outlook.GetNameSpace("MAPI")
        mail = outlook.CreateItem(0)
        # mail.SendUsingAccount = oacctouse
        if oacctouse:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))

        # mail.From = frmeml
        mail.To = eml
        mail.Subject = subj + str(emp_ID)
        body = "Good Day!!<br><br>Thank you for participating in the official corporate photoshoot!<br><br>Please find attached your photograph which you can use internally on C&Me and externally on LinkedIn only.<br><br><strong>It is mandatory to upload your photo on C&Me. Kindly complete this activity within 15 days of receipt of your photograph.</strong><br><br>Being an organization-sponsored shoot, your photograph is now the intellectual property of GBSI and will be kept on file till deemed necessary.<br><br>If you have any queries on the usage of this photograph, please contact the Internal Communications team by replying to this email.<br><br>Regards,<br><br>Internal Corporate Communications"
        mail.HTMLBody = (body)


        # if len(df_Content) !=0:
        #     nt = df_Content.to_string(index=False)
        #     nt = nt.replace("Mail_content","")
        #     mail.HTMLBody = "{0}</br></br><span style='font-size:12.0pt;background:yellow;mso-highlight:yellow'></span></br></br></br><p>Thank you.</p></br><p>Regards,<br/>Internal Corporate Communications</p>".format(df_Content.to_html(header=False,index=False,justify='left',border='0'),nt)
        # else:
        #     mail.HTMLBody = "{0}</br></br><p>Thank you.</p></br><p>Regards,</p>Internal Corporate Communications".format(df_Content.to_html(header=False,index=False,justify='left',border='0'))

        try:
            mail.Attachments.Add(os.path.join(os.getcwd(),str(emp_ID) + '.jpg'))
            df.loc[index,'Mail_Status'] = "Sent"
            # mail.display()                
            mail.Send()
        except:               
            df.loc[index,'Mail_Status'] = "File not found"       

    with pd.ExcelWriter(file.name,engine='openpyxl',mode='a',if_sheet_exists= 'replace') as writer:
        df.to_excel(writer,sheet_name="Mailing_list",index=False)

    messagebox.showinfo("Thank you!!","Completed..")



canvas = Canvas(width=550, height=230, bg='blue')
canvas.pack(expand=NO, fill=X)

image = ImageTk.PhotoImage(file="mLogo.jpg")
canvas.create_image(20, 20, image=image, anchor=NW)



pthLbl = Label(root,text='Choose File')
pthLbl.place(x=20,y=5)

btn = Button(root, text ='Select Emails file', command = lambda:open_file())
btn.place(x=450, y=50)

btnC = Button(root, text ='Convert_to_jpg', command = lambda:png_Jpg())
btnC.place(x=455, y=100)

btn_Mail = Button(root, text ='Mail_Dist.', command = lambda:Mail_dist())
btn_Mail.place(x=465, y=150)

  
mainloop()

# %%
