

#%%
import urllib.request
from bs4 import BeautifulSoup


import pandas as pd
from urllib.parse import urlparse,urljoin
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning

class Baltice_d():
    def Eurogate_D(self):
        # with urllib.request.urlopen('https://www.eurogate.de/eportal/state/show?_state=1vmjpsk8a1p3r&_unique=55jtppd4l1kw&_transition=start') as response:
        #         html = response.read().decode('utf-8')

        html = urllib.request.urlopen('https://www.eurogate.de/eportal/state/show?_state=1vmjpsk8a1p3r&_unique=55jtppd4l1kw&_transition=start').read()
        # htmlParser example : https://stackoverflow.com/questions/41687476/using-beautiful-soup-to-find-specific-class

        tbls = BeautifulSoup(html, "lxml").find_all('table', attrs={"class":"favorites"})



        requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

       # print(len(tbls))
        nDf = pd.DataFrame()
        url = 'https://www.eurogate.de'
        for tblN in tbls:
                for link in tblN.find_all('a', href=True):
                        if link['href'] .find('segelliste/state/do')!= -1:
                                full_path = urljoin(url , link['href'])
                                r = requests.get(full_path,verify=False)

                                # soup = BeautifulSoup(r.content,'html5lib')
                                soup = BeautifulSoup(r.content,"html.parser")
                                tbl  = soup.find('table',{'class':"rider"})
                                ##%%===================================================================
                                Links=list()
                                for lnk in tbl.find_all(href=True):
                                        Hlnk = 'https://www.eurogate.de' + lnk['href']
                                        H_r= requests.get(Hlnk,verify=False)

                                        H_soup = BeautifulSoup(H_r.content,"html.parser")
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
        nDf.to_excel(file_name,index=False,sheet_name='Baltic_d_eurogate')
        file_name.save()          

a = Baltice_d()
a.Eurogate_D()
