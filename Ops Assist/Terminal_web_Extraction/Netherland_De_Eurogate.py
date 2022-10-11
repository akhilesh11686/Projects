from dbm import ndbm
import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.packages.urllib3.exceptions import InsecureRequestWarning



class Netherland_de:
    def de_eurogate():
        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

        nDf = pd.DataFrame()

        headers = {"Accept-Language": "en-US,en;q=0.5"}

        url = 'https://www.eurogate.de/segelliste/state/show?_state=136707i70d5e6&_unique=1iubt6156s864&_transition=start&period=2&internal=false&languageNo=30&locationCode=HAM&order=%2B1'
        r= requests.get(url,verify=False,headers=headers)
        soup = BeautifulSoup(r.content,'html5lib')

        # print(soup.prettify())
        tbl  = soup.find('table',{'class':"rider"})
        ##%%===================================================================
        Links=list()
        for lnk in tbl.find_all(href=True):
            Hlnk = 'https://www.eurogate.de' + lnk['href']
            H_r= requests.get(Hlnk,verify=False)

            H_soup = BeautifulSoup(H_r.content,'html5lib')
            Rsl_tbl  = H_soup.find('table',{'class':"resultlist"})
            trs = Rsl_tbl.find_all('tr')

            rw_LIST =[]
            # TH
            for rw_H in trs[0:1]:
                td_H = rw_H.find_all('th')
                for td1 in td_H:
                    myString = td1.text
                    
                    if myString != "":
                        removal_list = [' ', '\t', '\n']
                        for s in removal_list:
                            myString = myString.replace(s, '')
                    rw_LIST.append(myString)
                nDf = nDf.append([rw_LIST])
                rw_LIST = []
            # td

            for rw in trs[2:]:
                tds = rw.find_all('td')
                for td in tds:
                    myString = td.text
                    removal_list = [' ', '\t', '\n']
                    for s in removal_list:
                        myString = myString.replace(s, '')            
                    rw_LIST.append(myString)
                nDf = nDf.append([rw_LIST])
                rw_LIST = []
        file_name =  pd.ExcelWriter('Terminals_Data.xlsx', engine='openpyxl',mode='a',if_sheet_exists='replace')
        nDf.to_excel(file_name,index=False,sheet_name='Netherland_De_Eurogate')
        file_name.save() 
                   
    # print('Done')