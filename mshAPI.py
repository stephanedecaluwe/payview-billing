import os
import datetime
import calendar
import re
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import csv
import logging
import codecs

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class BrowseBillingException(Exception):
    pass

class DownloadBillingAPI:
    def __init__(self, dossierStockageRés,numéroMois, année):
        _, nbDaysMoisFactu = calendar.monthrange(année, numéroMois)

        lastDatetimeMoisFactu = datetime.datetime(année,numéroMois, nbDaysMoisFactu, hour=23, minute=59, second=59, microsecond=999999)
        firstDateTimeMoisFactu = datetime.datetime(année, numéroMois,1,0,0,0,0)

        self.tsstart = int( firstDateTimeMoisFactu.timestamp()*1000)
        self.tsend = int( lastDatetimeMoisFactu.timestamp()*1000 )

        self.DossierStockageZipFiles = dossierStockageRés

        self.ClientsARecupToInfos = {}

        self._proxies = None #{"http":"http://85.115.60.150:80" ,"https":"http://85.115.60.150:80" }

        self._uri ='https://msh-portal-eu.icloud.ingenico.com'
        self._jwt = None

    def callAPI(self, endpoint,method,contentType="text/plain", jsonData=None, payload=None, queryString=None, stream=False):
        url = self._uri + endpoint
        jwt = self._jwt

        headers = {'cache-control': "no-cache",'content-type': contentType} #,'Connection':"keep-alive" }
        if jwt:
            headers['authorization'] = jwt

        logging.debug(f"Appel {method} on {endpoint}")
        response= requests.request(url = url,method=method, headers=headers,stream=stream, verify=False,params=queryString, proxies=self._proxies,json=jsonData,data=payload,
        timeout=40)
        # #Remove BOM !
        # if response.content[:3] == codecs.BOM_UTF8:
        #     response.content = response.content[3:]

        recJwt  = response.headers.get('authorization',None)
        if recJwt and not self._jwt:
            self._jwt = recJwt

        if (response.status_code in [200,203,204]  and response.headers['content-type'] == 'application/json'):
            logging.info(f'callAPI OK ({endpoint})')
            if(response.content):
                return response.json()

        elif (response.status_code == 200 and response.headers['content-type'] == 'text/csv; charset=UTF-8'):
            matched = re.match(r"^attachment; filename=(?P<filename>.*)$", response.headers['content-disposition'])
            if(not matched):
                raise BrowseBillingException(f"pas de match filename sur {response.headers['content-disposition']}" )

            fileNameRead = matched.group("filename")
            return fileNameRead, response.text
        else:
            logging.error(f"[{response.status_code}]")
            raise BrowseBillingException(f"callAPI KO {response.request.body}" )

    def login(self,user, password):
        self.callAPI('/v1/login',"POST","application/x-www-form-urlencoded",payload={'login':user,'password':password,'captcha':""})

    def loginAsIngenico(self):
        #Remplacer par login/mdp Laure PassPort
        self.login('stdecaluwe','123Soleil321%')  

    def logout(self):
        if(not self._jwt):
            return
        self.callAPI('/v1/logout','POST',"application/x-www-form-urlencoded")

    def getOffersList(self):
        items = self.callAPI('/v1/offer',"GET")
        return  {item['name']: item['id'] for item in items}

    def getClientsList(self):
        logging.info("------------[PassPort] Récup nom des clients à facturer par API")
        dicOffers = self.getOffersList()
        idAxisSimOnly = dicOffers.get("SIMs_Axis_only", None)

        self.ClientsARecupToInfos={}

        items = self.callAPI('/v1/customer',"GET","text/plain")

        for item in items:
            clientName = item['name'].strip()

            if('caspit' in clientName.lower()  or item['status'] != "Activated"):  #or 'progecarte' in  clientName.lower() 
                continue

            self.ClientsARecupToInfos[ clientName ] = {'ID': item['id'], 'simOnly': item.get('offerId',None) == idAxisSimOnly }
            logging.debug("{}->{}".format(clientName,item['id'] ))

    def downloadExcelDataFile(self, fileType, custID,clientName, formatFile='csv' ):
        if(not self._jwt):
            self.loginAsIngenico()

        params = {'customerId':custID, 'begin':self.tsstart ,'end':  self.tsend, 'format':formatFile  }

        fileread, content = self.callAPI(f'/v1/reporting/devices/activities/{fileType}','POST',"application/x-www-form-urlencoded",queryString=params )

        if len(content)<2:
            logging.info(f"[PassPort] fichier {fileType} {clientName} vide")
            return None

        fileName = f"{clientName}_{fileread}"
        csvFilePath = os.path.join(self.DossierStockageZipFiles,fileName  )

        with open( csvFilePath , 'w',encoding='utf-8') as cvFile:
            cvFile.write(content) #ecriture fichier

        logging.debug("File written OK " + csvFilePath)
        return fileName

    def RécupèreLesFichiersClientsAFacturerAsIngenico(self,ignoredListMinuscule):
        for nomClient in sorted(self.ClientsARecupToInfos ):
            if 'sandbox' in nomClient.lower() or (ignoredListMinuscule and nomClient.lower() in ignoredListMinuscule):
                logging.debug(f"PassPort ignore récup {nomClient}")
                continue

            logging.info("[PassPort] Récup billing data pour " + nomClient)

            IDclient = self.ClientsARecupToInfos[nomClient]['ID']
            
            if self.ClientsARecupToInfos[nomClient]['simOnly'] :
                logging.info("[PassPort] Terminal skippe  car simOnly sur " + nomClient)
            else:
                self.downloadExcelDataFile("poi-connections",IDclient,nomClient ) 
            
            self.downloadExcelDataFile( "sim-status" , IDclient,nomClient)

    def RécupèreFichiersFactu(self,ignoredListMinuscule):
        self.loginAsIngenico()
        self.getClientsList()
        self.RécupèreLesFichiersClientsAFacturerAsIngenico(ignoredListMinuscule)

