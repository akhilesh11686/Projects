#%%
import datetime
import getpass
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.font import Font
from PIL import ImageTk, Image
import os

from CMA_CGM_data import getCMA
from mainFile import mainProgram

# Loop Extract of Partner vessels
def getExt():
    if varMaerks.get() ==1:
        mainProgram.fnd(xlPth,'Maersk')

    if varCosco.get()==1:
        mainProgram.fnd(xlPth,'Cosco')
        # pass

    if varHapag.get()==1:
        mainProgram.fnd(xlPth,'Hapag')

    if varMsc.get()==1:
        mainProgram.fnd(xlPth,'Msc')


def myupdate():
    if varAll.get()==1:
        chkMaerks.select()
        chkCosco.select()
        chkHpag.select()
        chkMsc.select()
    else:
        chkMaerks.deselect()
        chkCosco.deselect()
        chkHpag.deselect()
        chkMsc.deselect()    

def Extraction_Cma():
    getCMA.callData()

def ChooseFile():
    global xlPth     
    xlPth = filedialog.askopenfilename(initialdir = os.getcwd(),title= "select file",filetypes = (("Excel file","*.xl*"),("All files","*.*")))

    if len(xlPth)==0:
        messagebox.showerror('error', 'Something went wrong!')
    else:        
# Label path
        entrBar =ttk.Label(frmLblLeft,text= xlPth)
        entrBar.pack()
        entrBar.place(x=120,y=60,width=240,height=25)

        frmPtnrChoose = ttk.Labelframe(frameLeft,text="Partners Web Extract")
        frmPtnrChoose.pack(side="top",fill='both',expand="yes")
        frmPtnrChoose.pack()

        global chkAll, chkMaerks,chkHpag,chkMsc,chkCosco,varAll,listChck

        global varMaerks,varCosco,varHapag,varMsc

        varMaerks = IntVar(frmPtnrChoose)
        varCosco = IntVar(frmPtnrChoose)
        varHapag = IntVar(frmPtnrChoose)
        varMsc = IntVar(frmPtnrChoose)
        
        
        chkMaerks = Checkbutton(frmPtnrChoose, text="Maersk", variable=varMaerks)
        chkMaerks.grid(row=1, column=1)

        chkCosco = Checkbutton(frmPtnrChoose, text="Cosco", variable=varCosco)
        chkCosco.grid(row=1,column=2)

        chkHpag = Checkbutton(frmPtnrChoose, text="Hapag Lyod", variable=varHapag)
        chkHpag.grid(row=1,column=3)

        chkMsc = Checkbutton(frmPtnrChoose, text="MSC", variable=varMsc)
        chkMsc.grid(row=2,column=3)        

        varAll = IntVar(frmPtnrChoose)
        chkAll = Checkbutton(frmPtnrChoose,text="Select All",variable=varAll,command=myupdate)
        chkAll.grid(row=1,column=4)
        
        exBtn = Button(frmPtnrChoose, text="Extraction",command=getExt)
        exBtn.grid(row=1,column=7,sticky='ne')





        
# Time and date  ------------------------------------
def quit(*args):
    LeftLableFrame.destroy()

def clock_time():
    time= datetime.datetime.now()
    time= (time.strftime("Date: %Y -%m -%d, Time: %H:%M:%S"))
    txt.set(time)
    LeftLableFrame.after(1000,clock_time)

 

win =Tk()
win.title("Terminal | Parternes Web Extraction")
# Set the size of the tkinter window
win.geometry("950x600")
win.maxsize(950, 600) # specify the max size the window can expand to
win.minsize(950,600)
win.config(bg="#09669f")



# Create an instance of ttk style
s = ttk.Style()
s.theme_use('default')
s.configure('TNotebook.Tab', background="White")
s.map("TNotebook", background= [("selected", "gray")])

# Create a Notebook widget
nb = ttk.Notebook(win)

# Add a frame for adding a new tab
Frm1= ttk.Frame(nb, width= 850, height=300)
nb.add(Frm1, text= 'Single')


LeftLableFrame = ttk.Labelframe(Frm1,text="Web Extr")
LeftLableFrame.pack(side="top",fill='both',expand="yes")


# Login user ----------------------------------------
lblLogin= Label(LeftLableFrame,text='Welcome :'+" "+ getpass.getuser(),font=("arial italic", 14,),anchor='w')
lblLogin.pack()
lblLogin.place(x=620,y=0,width=280,height=18)


LeftLableFrame.after(1000,clock_time)

fnt = Font(family = "helvetica",size=10, weight ="bold")
txt = StringVar()
lblTime = Label(LeftLableFrame,textvariable=txt, font = fnt, foreground= "black")
lblTime.place(x=620,y=30,width=220,height=30)


# Picture -----------------------------------------------------------
frame = Frame(LeftLableFrame, width=900, height=515)
frame.pack()
frame.place(x=420,y=80,width=480,height=360)

canv = Canvas(master=frame)
canv.place(x=0, y=0, width=480, height=360)

pth = os.getcwd() + "\\cgmLog.jpg"
# pth = os.getcwd() + "\\cgmLog.jpg"
image = Image.open(pth)

image = image.resize((480, 360))
image = ImageTk.PhotoImage(image)
label = Label(frame, image = image)
canv.create_image(0, 0, image=image, anchor='nw')


# 'Left labelFrame'-------------------------------------------------
frameLeft = Frame(LeftLableFrame, width=400, height=400)
frameLeft.pack()
frameLeft.place(x=10,y=80,width=400,height=360)

frmLblLeft = ttk.Labelframe(frameLeft,text="Partners Web Data")
frmLblLeft.pack(side="top",fill='both',expand="yes")

xlPth = ""   #Varible





cmaBtn =ttk.Button(frmLblLeft, text ="CMA Web Data", command = Extraction_Cma)
cmaBtn.pack()
cmaBtn.place(x=5,y=20,width=380,height=25)


# Get path
PthBtn =ttk.Button(frmLblLeft, text ="Choose file", command = ChooseFile)
PthBtn.pack()
PthBtn.place(x=5,y=60,width=80,height=25)










# ==================2 Tabs===========================================================
                    # Multiple Option #
# ==================2 Tabs===========================================================

def MultSelect():
    
    if varT.get()==1:
        chkTA.select()
        chkTB.select()
        chkTC.select()
        chkTD.select()
        chkTE.select()
        chkTF.select()
        chkTG.select()
        chkTH.select()
        chkTI.select()                
        chkTJ.select()
        chkTK.select()
        chkTL.select()        
    else:
        chkTA.deselect()
        chkTB.deselect()
        chkTC.deselect()
        chkTD.deselect()
        chkTE.deselect()
        chkTF.deselect()
        chkTG.deselect()
        chkTH.deselect()
        chkTI.deselect()                
        chkTJ.deselect()
        chkTK.deselect()
        chkTL.deselect()     

def getTerminal_Ext():
    if varA.get() ==1:
        mainProgram.get_terminals('Antwerp-GATEWAY 1700')

    if varB.get()==1:
        mainProgram.get_terminals('Antwerp-PSA Q913')

    if varC.get()==1:
        mainProgram.get_terminals('Antwerp-PSA TERMINAL 869')

    if varD.get() ==1:
        mainProgram.get_terminals('Netherland DE: Eurogate')

    if varE.get()==1:
        mainProgram.get_terminals('Rotterdam-World_Gateway')

    if varF.get()==1:
        mainProgram.get_terminals('BREMERHAVEN CMA3PF')

    if varG.get() ==1:
        mainProgram.get_terminals('LUBECK INTRA_BALT')

    if varH.get()==1:
        mainProgram.get_terminals('DpWorld_London')

    if varI.get()==1:
        mainProgram.get_terminals('DpWorld_Southampton')

    if varJ.get() ==1:
        mainProgram.get_terminals('port_Of_Felixstowe')

    if varK.get()==1:
        mainProgram.get_terminals('Baltic_E_eurogate')

    if varL.get()==1:
        mainProgram.get_terminals('Baltic_D_eurogate')



Frm2 = ttk.Frame(nb, width= 850, height=300)
nb.add(Frm2, text= "Multiple")
nb.pack(expand= True, fill=BOTH, padx= 10, pady=10)

multFrame = ttk.Labelframe(Frm2,text="Terminals Web Extraction")
multFrame.pack(side="top",fill='both',expand="yes")

# CMA 
cmaBtn =ttk.Button(multFrame, text ="CMA Web Data", command = Extraction_Cma)
cmaBtn.pack()
cmaBtn.place(x=10,y=20,width=400,height=25)

# Picture
frame1 = Frame(multFrame, width=900, height=515)
frame1.pack()
frame1.place(x=420,y=80,width=480,height=360)

canv = Canvas(master=frame1)
canv.place(x=0, y=0, width=480, height=360)

pth1 = os.getcwd() + "\\vslImag.JPEG"
# pth1 = os.getcwd() + "\\vslImag.JPEG"

image1 = Image.open(pth1)

image1 = image1.resize((480, 360))
image1 = ImageTk.PhotoImage(image1)
label = Label(frame, image = image1)
canv.create_image(0, 0, image=image1, anchor='nw')

frmLblLeft1 = ttk.Labelframe(multFrame,text="Terminals Web Data")
frmLblLeft1.pack()
frmLblLeft1.place(x=5,y=80,width=410,height=360)

global varA,varB,varC,varD,varE,varF,varG,varH,varI,varT,chkT_All

varA = IntVar(frmLblLeft1)
varB = IntVar(frmLblLeft1)
varC = IntVar(frmLblLeft1)

varD = IntVar(frmLblLeft1)
varE = IntVar(frmLblLeft1)
varF = IntVar(frmLblLeft1)

varG = IntVar(frmLblLeft1)
varH = IntVar(frmLblLeft1)
varI = IntVar(frmLblLeft1)

varJ = IntVar(frmLblLeft1)
varK = IntVar(frmLblLeft1)
varL = IntVar(frmLblLeft1)

varT = IntVar(frmLblLeft1)




chkT_All = Checkbutton(frmLblLeft1,text="Select All",variable=varT,command=MultSelect)
chkT_All.grid(row=2,column=1)

exBtn = Button(frmLblLeft1, text="Extraction",command=getTerminal_Ext)
exBtn.grid(row=2,column=4,sticky='ne')



chkTA = Checkbutton(frmLblLeft1, text="Antwerp_1700", variable=varA)
chkTA.grid(row=6, column=1)

chkTB = Checkbutton(frmLblLeft1, text="Antwerp_PSA_Q913", variable=varB)
chkTB.grid(row=8,column=1)

chkTC = Checkbutton(frmLblLeft1, text="Antwerp_PSA_869", variable=varC)
chkTC.grid(row=10,column=1)



chkTD = Checkbutton(frmLblLeft1, text="Netherland DE: Eurogate", variable=varD)
chkTD.grid(row=6, column=2)

chkTE = Checkbutton(frmLblLeft1, text="Rotterdam-World_Gateway", variable=varE)
chkTE.grid(row=8,column=2)

chkTF = Checkbutton(frmLblLeft1, text="BREMERHAVEN CMA3PF", variable=varF)
chkTF.grid(row=10,column=2)


chkTG = Checkbutton(frmLblLeft1, text="LUBECK INTRA_BALT", variable=varG)
chkTG.grid(row=12, column=1)

chkTH = Checkbutton(frmLblLeft1, text="DpWorld_London", variable=varH)
chkTH.grid(row=14,column=1)

chkTI = Checkbutton(frmLblLeft1, text="DpWorld_Southampton", variable=varI)
chkTI.grid(row=16,column=1)


chkTJ = Checkbutton(frmLblLeft1, text="port_Of_Felixstowe", variable=varJ)
chkTJ.grid(row=12, column=2)

chkTK = Checkbutton(frmLblLeft1, text="Baltic_E_eurogate", variable=varK)
chkTK.grid(row=14,column=2)

chkTL = Checkbutton(frmLblLeft1, text="Baltic_D_eurogate", variable=varL)
chkTL.grid(row=16,column=2)


win.mainloop()




#%%

# import pandas as pd
# df = pd.read_excel('Terminals_Data.xlsx',sheet_name='Cosco')
# df1 = pd.read_excel('Terminals_Data.xlsx',sheet_name='Sheet1')
# #%%
# # [print (idx1['Age']) for rw,idx in df1.iterrows() for rw1,idx1 in df.iterrows() if (idx['Name']==idx1['Name']) & (idx['SurName']==idx1['SurName'])]
# # df1['Age'] = [idx1['Age'] for rw,idx in df1.iterrows() for rw1,idx1 in df.iterrows() if (idx['Name']==idx1['Name']) & (idx['SurName']==idx1['SurName'])]
# df1['Age'] = [idx['Age'] for rw1,idx1 in df1.iterrows() for rw,idx in df.iterrows() if (idx1['Name']==idx['Name']) & (idx1['SurName']==idx['SurName'])]
# # df1['Age'] = df1.apply(lambda x: df['Age'] if(x['Name']==df['Name'] & x['SurName']==df['SurName']) else '')
