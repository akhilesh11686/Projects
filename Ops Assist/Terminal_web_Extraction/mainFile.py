#%%
# from _typeshed import self
import os
from re import M
import sys
from tkinter import messagebox

from numpy import nan
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

from maersk_web import Mrsk_web
from cosco_web import Cosco_web_e
from hapag_web import hapag_l_web
from msc_web import MSC_web

from Antwerp_869 import T_869
from Antwerp_913 import T_913
from Antwerp_1700 import T_1700
# from baltic_d_eurogate import Baltice_d   #bug exists
# from baltic_e_eurogate import Baltice_e   #bug exists
from DpWorld_London import dpLondon
from DpWorld_Southampton import dpSouthampton
from Germany_CMA3PF import BREMERHAVEN_CMA3PF
from Germany_INTRA_BALT import INTRA_BALT
from Netherland_De_Eurogate import Netherland_de
from ROTTERDAM_RWG import RWG_terminal
# from SHEKOU import shekou_terminals
from portOfFelixstowe import uk_Felixstowe


import xlsxwriter
#%%
class mainProgram:
    def fnd(pth,carrNm):

        df = pd.read_excel(r"{}".format(pth),sheet_name="vessel & service data")
        # initialize a empty df
        appended_data = pd.DataFrame()

        # if carrNm.find('MAERSK'):
        if 'Maersk' in carrNm.title():
            print('Maersk')
            maersk_df = df.loc[df['Operator'] == 'Maersk']
            for index, row in maersk_df.iterrows():
                vslNm = row['Vessel Name']
                vsl = row['CODE']
                
                if isinstance(vsl,float):
                    vsl = int(vsl)

                frmDate = row['fromDate'].strftime("%Y-%m-%d")
                toDate = row['toDate'].strftime("%Y-%m-%d")
                serverData = Mrsk_web.getWebData(vslNm,vsl,frmDate)
                # appened it
                appended_data = appended_data.append(serverData)

            file_name = pd.ExcelWriter('Partener_Vessels.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
            appended_data.to_excel(file_name,index=False,sheet_name='Maersk')
            file_name.save()           


        # Cosco web
        elif 'Cosco' in carrNm.title():
            print('cosco')
            Cosco_df = df.loc[df['Operator'] == 'Cosco']
            for index, row in Cosco_df.iterrows():
                vslNm = row['Vessel Name']
                vsl = row['CODE']
                
                if isinstance(vsl,float):
                    vsl = int(vsl)

                frmDate = row['fromDate']
                # toDate = row['toDate'].strftime("%Y-%m-%d")
                serverData = Cosco_web_e.cosco_getWebData(vslNm,vsl,frmDate)
                # appened it
                # appended_data = appended_data.append(serverData)  
                appended_data = pd.concat([appended_data,serverData])
                # messagebox.showinfo('Done','Thank you!')

            file_name = pd.ExcelWriter('Partener_Vessels.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
            appended_data.to_excel(file_name,index=False,sheet_name='Cosco')
            file_name.save()       


        # Hapag web
        elif 'Hapag' in carrNm.title():
            print('hapga')
            # Hapage_df = df.loc[df['Operator'] == 'Hapag']
            # for index, row in Hapage_df.iterrows():
            #     vslNm = row['Vessel Name']
            #     serverData = hapag_l_web.getHapag(vslNm)
            #     # appened it
            #     appended_data = appended_data.append(serverData)  
            # hapag_l_web.getHapag()   Olde version issuee ...............................

        # MSC web
        elif 'Msc' in carrNm.title():
            print('msc')
            Msc_df = df.loc[df['Operator'] == 'MSC']

            
            msc_list = Msc_df['Vessel Name'].to_list()
            serverData = MSC_web.getMSC(msc_list)

            # for index, row in Msc_df.iterrows():
            #     vslNm = row['Vessel Name']
            #     serverData = MSC_web.getMSC(vslNm)
            #     # appened it
            #     appended_data = appended_data.append(serverData)  

            # file_name = pd.ExcelWriter('Partener_Vessels.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
            # appended_data.to_excel(file_name,index=False,sheet_name='Msc')
            # file_name.save()               


    def get_terminals(myTerminals):
        # LoadingGif.startAnimation
        if 'Antwerp-PSA TERMINAL 869' in myTerminals:
            T_869.ant_t869()
            return
        elif 'Antwerp-PSA Q913' in myTerminals:
            T_913.ant_t913()
            return
        elif 'Antwerp-GATEWAY 1700' in myTerminals:
            T_1700.ant_t1700()
            return
        elif 'CMA3PF' in myTerminals:
            BREMERHAVEN_CMA3PF.BREMERHAVEN_t3PF()
            return
        elif 'INTRA_BALT' in myTerminals:
            INTRA_BALT.Germany_BALT()
            return
        elif 'World_Gateway' in myTerminals:
            RWG_terminal.Rotterdam_RWGt()
            return
        elif 'Eurogate' in myTerminals:
            Netherland_de.de_eurogate()
            return
        elif 'DpWorld_London' in myTerminals:
            dpLondon.london_Gateway()
            return
        elif 'DpWorld_Southampton' in myTerminals:
            dpSouthampton.Southampton_Gateway()
            return       
        elif 'port_Of_Felixstowe' in myTerminals:
            uk_Felixstowe.terminal_Felixstowe()
            return
        # elif 'Baltic_E_eurogate' in myTerminals:
        #     Baltice_e.Eurogate_E()
        #     return
        # elif 'Baltic_D_eurogate' in myTerminals:
        #     Baltice_d.Eurogate_D()
        #     return      
        # elif 'Shekou_terminal' in myTerminals:
        #     shekou_terminals.getShekou_e()
        #     return                