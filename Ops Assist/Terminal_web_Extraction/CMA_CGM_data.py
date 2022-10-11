
#%%
import xlwings as xw
from tkinter import messagebox

class getCMA():
	def callData():
		wb = xw.Book("Cma-Cgm_Data.xlsm")
		app = wb.app

		macro_vba = app.macro("'Cma-Cgm_Data.xlsm'!Clear_Sheet")
		macro_vba()

		macro_vba = app.macro("'Cma-Cgm_Data.xlsm'!SQL_Query")
		macro_vba()
		wb.save()
		wb.close()
		messagebox.showinfo("Processing..","Completed!!")
