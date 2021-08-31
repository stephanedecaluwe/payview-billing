import requests
import pprint
import tablib
import os
import sys
import re
import logging

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

TOKEN = '38e77df1-0b9e-4c2e-a00a-91e502f87ac0'
ROOT_URL = 'https://admin.payview.fr/api/'
PROXIES = None
HEADERS = {'Authorization':'Bearer '+ TOKEN }
REGEX_FILE = re.compile(r'^attachment; filename="(?P<client>.*)_(?P<annee>\d{4})-(?P<mois>\d{2})_(?P<fileType>.*).csv"$')
fileTypeName = {'poi-connections':'terminal_connections','sim-status':"sim_status"}

def exportListeVersExcel(filePath, liste, titre="export"):
    if not liste:
        logging.error(f"Liste vide pour export {titre}")
        return

    tabCli = tablib.Dataset(title=titre, headers= liste[0].keys() )

    for l in liste:
        tabCli.append( l.values() )

    with open( filePath, mode='wb') as f:  #PermissionError si déjà ouvert
            f.write(tabCli.export('xlsx'))

def callAPI(uri, query=None):
    print(ROOT_URL+uri, query)
    res= requests.request("GET",url=ROOT_URL+uri, headers= HEADERS,params=query, verify=False,proxies=PROXIES ).json()
    print("res " ,res)
    return res

def getCustomers(size=200,offset=0):
    return callAPI('customers', query={'size':size,'offset':offset})

"""Retourne la liste des clients finaux (pas sous les grossistes)"""
def getAllCustomers():
    offset = 0
    pageSize = 200
    count = 0
    allCustomers = []

    while count ==0:
        res = getCustomers(pageSize, offset)
        tab = res['items'] #{'clientName': 'ADIDAS FRANCE', 'grossisteName': 'INGENICO', 'activated': True, 'simOnly': False, 'idClient': 'b5954a19-4ef5-456e-b858-6611b23bd1fd'}
        count = pageSize-len(tab)
        offset += pageSize

        allCustomers += tab

    return allCustomers

def getFilesThisCustomer(clientObj,numMois, année, dossierCsv , writeFile=True):
    cid     = clientObj['idClient']
    cname   = clientObj['clientName']
    
    if not clientObj['activated']:
        logging.warning(f'client {cname} désactivé')
        return

    logging.info(f"Récup fichiers {cname}")
   
    
    getFactuFile('poi-connections', cid, numMois, année, dossierCsv , writeFile)
    #if not clientObj['simOnly']:
    getFactuFile('sim-status', cid, numMois, année, dossierCsv , writeFile)
    
# fileType: 'poi-connections' 'sim-status'
def getFactuFile(fileType, clientId, numMois, année, dossierCsv , writeFile=True):
    def getFile(url):
        encodingUsed = 'utf_16_le'
        response = requests.request("GET",url=url, verify=False,proxies=PROXIES )
        response.encoding = encodingUsed

        h = response.headers
        if h['Content-Type'] != 'text/csv; charset=utf-8':
            logging.error(f"Pas le bon format de retour de fichier {h['Content-Type']} {url}")
            #raise Exception("KO")
            return      
        m = REGEX_FILE.match(h['Content-Disposition']) #'attachment; filename="ACCES VITAL TECHNOLOGY_2020-10_terminal_connections.csv"'
        if not m:
            logging.error(f"Pas de matching sur {h['Content-Disposition']}")
            return
        assert int(m.group('annee')) == année and int(m.group('mois'))==numMois, f"Erreur année, mois lues, {m.group('annee')} {m.group('mois')} {h['Content-Disposition']}"
        assert m.group('fileType')== fileTypeName[fileType], f"Mauvais fileType '{m.group('fileType')}'"

        print( m.group('client'), m.group('fileType') )
        if not writeFile:
            return
        
        filename = f"{m.group('client')}_{fileType}.csv"
        csvFilePath = os.path.join(dossierCsv,filename  )

        with open( csvFilePath , 'w',encoding=encodingUsed) as csvFile:
            csvFile.write(response.text) #ecriture fichier
        return filename

    res = callAPI(f"activities/{fileType}", query={'date':f'{année}-{numMois:02d}','clientId':clientId,'format':'csv'})

    return getFile( res['url'] )


if __name__ == "__main__":
    dossierRes = r'C:\Users\squack\Documents\_PAYVIEW\FACTURATION\Generated\TMP'

    tab= getAllCustomers()
    exportListeVersExcel(os.path.join(dossierRes,"clients.xlsx"), tab, titre="Clients")
    mois = 10
    année = 2020

    for c in tab:
        getFilesThisCustomer(c,mois, année, dossierRes , writeFile=True)



    




