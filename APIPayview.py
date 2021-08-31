import requests
import pprint
import tablib
import os
import sys
import colorlog
import time
import logging
import re
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class APIPayview:
    def __init__(self, ssdomain, login, mdp):
        self.ssdomain = ssdomain
        self.login = login
        self.mdp = mdp

        self._cookies = None
        self.proxies = None #{"http":"http://85.115.60.150:80" ,"https":"http://85.115.60.150:80" }
        
        # LOGIN AS
        self._ssdomainAs = None
        self._cookiesAs = None

        # General
        self._orgaUUID = None
        self.URI = f'https://{self.getSsdomain()}.payview.fr/api'
        self.HEADERS = {
                'Connection':"keep-alive",
                'accept': 'application/json',
                #'content-type': 'application/json;charset=UTF-8',
                'referer':f'https://{self.getSsdomain()}.payview.fr/',
                'Host': f'{self.getSsdomain()}.payview.fr'
            }

    def getCookie(self):
        if self._cookiesAs:
            return self._cookiesAs
        return self._cookies

    def getSsdomain(self):
        if self._ssdomainAs:
            return self._ssdomainAs
        return self.ssdomain

    def callAPI(self, method,uri, jsonData=None,query=None):
        return requests.request(method,url=self.URI+uri, headers=self.HEADERS,json=jsonData,params=query, verify=False,proxies=self.proxies,cookies=self.getCookie(),timeout=40 )

    def Login(self):
        res = self.callAPI("POST", '/login', jsonData={'email':self.login,'password':self.mdp})
        #res = requests.post(self.URI+'/login',headers=self.HEADERS,json={'email':self.login,'password':self.mdp},verify=False,proxies=None )
        self._cookies = res.cookies

    def LoginAs(self,token):
        res = self.callAPI("POST","/login", jsonData={'token': token})
        self._cookiesAs = res.cookies
    
    def LogoutAs(self):
        self._cookiesAs = self._ssdomainAs = None

    def getSession(self):
        res = self.callAPI("GET", '/session' ).json()
        self._orgaUUID = res['organization']

    def getTerminals(self, pageSize=50, offset=0,SNSearched=None,fournisseur=None ):
        if not self._orgaUUID:
            self.getSession()
        params = {'provider':self._orgaUUID,'limit':pageSize,'offset':offset }
        if SNSearched:
            params['posTerminal.serialNumber']= SNSearched
        if fournisseur:
            params['directProvider.legalName'] = fournisseur
        
        res = self.callAPI("GET", '/posTerminalSubscriptions',query=params ).json()

        return res

    def getSims(self,pageSize=20, offset=0,iccid=None,nomOrga=None ):
        if not self._orgaUUID:
            self.getSession()
        params = {'provider':self._orgaUUID,'limit':pageSize,'offset':offset }

        if iccid:
            params['sim.iccid'] = iccid
        if nomOrga:
            params['client.legalName'] = nomOrga

        res = self.callAPI("GET", '/simSubscriptions',query=params ).json()

        return res

        # {"items": [
		# {
		# 	"uuid": "27d1168e-ebe6-4724-8896-be78cc3ad331",
		# 	"createdAt": "2020-09-24T10:17:28.838164+00:00",
		# 	"updatedAt": "2020-11-10T10:20:41.948719+00:00",
		# 	"provider": "4ac85a69-8e6e-4f8f-adbb-4c62aba60dd6",
		# 	"client": {
		# 		"uuid": "860e553d-6328-4b04-8d0c-ccf3d764cf18",
		# 		"provider": "4ac85a69-8e6e-4f8f-adbb-4c62aba60dd6",
		# 		"legalName": "AVT"
		# 	},
		# 	"sim": {
		# 		"uuid": "a3b6e308-80da-4cff-9218-231b80cd1872",
		# 		"createdAt": "2020-09-18T17:01:18.391435+00:00",
		# 		"updatedAt": "2020-11-10T10:20:41.948719+00:00",
		# 		"label": " ",
		# 		"iccid": "89332401000015992403",
		# 		"status": "PENDING_ACTIVATION",
		# 		"tariffData": null,
		# 		"tariffDataUnit": null,
		# 		"activationDate": null,
		# 		"endOfEngagementDate": null,
		# 		"lastCommunicationDate": null,
		# 		"firstCommunicationDate": null,
		# 		"posTerminal": null,
		# 		"dataUsage": null
		# 	}
		# }],"totalResults": 48891}

    def resilierSim(self, uuidList ):
        if not isinstance(uuidList,list):
            uuidList = [uuidList]

        res = self.callAPI("POST","/sims/terminate", jsonData={'uuids':uuidList }).json()
        return res #{}

    def getOrganizations(self,pageSize=20, offset=0, nomOrgaSearched =None, orgaUuid=None,typeOrga="WHITE_LABEL" ):
        params = {'limit':pageSize,'offset':offset,'type':typeOrga }

        if nomOrgaSearched:
            params['legalName'] = nomOrgaSearched

        if orgaUuid == None:
            if not self._orgaUUID:
                self.getSession()
            orgaUuid = self._orgaUUID
            
        logging.info(f"[PayView] Lecture organisations de {orgaUuid} offset={offset}")
        res = self.callAPI("GET", f'/organizations/{orgaUuid}/clients',query=params ).json()
        return res
    #{'items':[{businessName: null,createdAt: "2020-06-15T15:52:04.132001+00:00",legalName: "GRAND FRAIS (FUJITSU)",registrationNumber: "38780672200022",type: "WHITE_LABEL"
    #updatedAt: "2020-11-12T13:12:39.847317+00:00",uuid: "cf9c18bb-aadb-4629-b258-0b5e54850138"}],
    #totalResults: 1}

    def getAllClients(self, orgaUuid=None, grossisteList=['HM TELECOM','AVT','SATIN','IPSF','BRED BANQUE POPULAIRE','LM CONTROL'], parentGrossite=None):
        pageSize =20
        offset = 0
        tabRes = []

        more = True

        while more:
            res = self.getOrganizations(pageSize=pageSize, offset =offset,orgaUuid=orgaUuid)['items']

            for r in res:
                legalName = r['legalName'] 
                if grossisteList and legalName in grossisteList:
                    tabRes += self.getAllClients( r['uuid'] , grossisteList=None,parentGrossite=legalName )
                else:
                    tabRes.append( {'legalName': legalName, 'uuid':r['uuid'], 'grossiste':parentGrossite} )

            more = len(res) > 0
            offset += pageSize

        return tabRes

    def getUsers(self,pageSize=20, offset=0,organizationId = None, emailSearched=None):
        if not self._orgaUUID:
            self.getSession()

        params = {}
        if emailSearched:
            params['email'] = emailSearched
        
        if organizationId:
            params['organization'] = organizationId
        else:
            params['organization'] = self._orgaUUID

        params['limit'] = pageSize
        params['offset'] = offset

        res = self.callAPI("GET", f'/users?email={emailSearched}&organization={organizationId}&limit={pageSize}&offset={offset}')

        return res.json()

    def creerCompteAcces(self, email, labelCompte, contratsUidList):
        params = {'provider':self._orgaUUID, 'email':email, 'label':labelCompte, 'contractProfiles':[{'contract': uid} for uid in contratsUidList]}
        res = self.callAPI("POST","/contractsAccesses",jsonData=params ).json()
        assert res['success']

    def getComptesAcces(self,pageSize=50, offset=0,emailRecherche=None):
        if not self._orgaUUID:
            self.getSession()

        params = { 'limit':pageSize,'offset':offset }
        if emailRecherche:
            params['email'] = emailRecherche
        res = self.callAPI("GET", f'/organizations/{self._orgaUUID}/contractsAccesses',query=params ).json()
        #res = requests.get(self.URI+f'/organizations/{self._orgaUUID}/contractsAccesses', params=params, headers=self.HEADERS, cookies=self._cookies).json()
        if not res['items']:
            return None
        else:
            return res['items']

        #pprint.pprint(res.json())
        #{'items': [{'createdAt': '2020-10-26T07:58:30.325258+00:00',
            # 'email': 'thierry.cade@ingenico.com',
            # 'label': 'Test Portail COM TCA',
            # 'provider': 'fb10e478-7981-4cd7-8fb0-d980e7412dc7',
            # 'status': 'ACTIVE',
            # 'updatedAt': '2020-10-26T07:58:30.325258+00:00',
            # 'uuid': '55146a48-b180-459a-bd26-a4bb6b65a80e'}..],'totalResults': 6}

    def detailsCompteDacces(self, compteUUID,pageSize=50, offset=0):
        params = { 'limit':pageSize,'offset':offset }
        res = self.callAPI("GET", f'/contractsAccesses/{compteUUID}',query=params ).json()

        return {'contracts': res.get("contractProfiles",None), 'users': res['users'],"uuid":res["uuid"]}
        # {"uuid": "55146a48-b180-459a-bd26-a4bb6b65a80e",
        #  "createdAt": "2020-10-26T07:58:30.325258+00:00", 
        #  "updatedAt": "2020-10-26T07:58:30.325258+00:00", 
        #  "label": "Test Portail COM TCA", 
        #  "provider": "fb10e478-7981-4cd7-8fb0-d980e7412dc7", 
        #  "status": "ACTIVE", "contractProfiles": [
        #      {"uuid": "c9ca5b91-7c40-4130-82b4-18b8fb924ee1", "createdAt": "2020-10-26T07:58:30.334008+00:00", "updatedAt": "2020-10-26T07:58:30.334008+00:00", "application": "CBEMV", "bankCode": "30003", "rank": null, "contract": "3619346", "merchantLabel": "COM TCA"}, 
        #      {"uuid": "268c528a-7fb6-46fd-99c4-141d17d2d235", "createdAt": "2020-10-26T07:58:30.338473+00:00", "updatedAt": "2020-10-26T07:58:30.338473+00:00", "application": "CBCLESS", "bankCode": "26550", "rank": null, "contract": "1999281", "merchantLabel": "COM 1"}, 
        #      {"uuid": "4c845fba-810c-426d-8299-2d8c2badf57f", "createdAt": "2020-10-26T07:58:30.342159+00:00", "updatedAt": "2020-10-26T07:58:30.342159+00:00", "application": "CBCLESS", "bankCode": "30004", "rank": null, "contract": "4338295", "merchantLabel": "COM 2"}, 
        #      {"uuid": "11772efa-28e1-401b-b080-dd61e7d75a06", "createdAt": "2020-10-26T07:58:30.352955+00:00", "updatedAt": "2020-10-26T07:58:30.352955+00:00", "application": "CBEMV", "bankCode": "30004", "rank": null, "contract": "4338292", "merchantLabel": "COM 3"}, 
        #      {"uuid": "d3cc4d4e-9556-418d-adc8-c25c16403fb4", "createdAt": "2020-10-26T07:58:30.356316+00:00", "updatedAt": "2020-10-26T07:58:30.356316+00:00", "application": "CBEMV", "bankCode": "30004", "rank": null, "contract": "4338295", "merchantLabel": "COM 4"}, 
        #      {"uuid": "700596ae-3562-4c4c-9115-dc964bacfeff", "createdAt": "2020-10-26T07:58:30.360142+00:00", "updatedAt": "2020-10-26T07:58:30.360142+00:00", "application": "CBCLESS", "bankCode": "30001", "rank": null, "contract": "2330301", "merchantLabel": "COM 5"}, 
        #      {"uuid": "dd886d02-2b8c-4499-8eba-e9494646c0d5", "createdAt": "2020-10-26T07:58:30.374195+00:00", "updatedAt": "2020-10-26T07:58:30.374195+00:00", "application": "CONECS", "bankCode": "10000", "rank": null, "contract": "1207847", "merchantLabel": null}, 
        #      {"uuid": "9e72f3a8-beb0-4e90-8d15-cf34b7383960", "createdAt": "2020-10-26T07:58:30.380033+00:00", "updatedAt": "2020-10-26T07:58:30.380033+00:00", "application": "CBCLESS", "bankCode": "11899", "rank": null, "contract": "4278622", "merchantLabel": null}, 
        #      {"uuid": "576da3f2-fda2-4b82-b918-8dab0e84ded1", "createdAt": "2020-10-26T07:58:30.392022+00:00", "updatedAt": "2020-10-26T07:58:30.392022+00:00", "application": "CBCLESS", "bankCode": "11899", "rank": null, "contract": "4278155", "merchantLabel": null}, 
        #      {"uuid": "9fbae29c-c1b7-4f9b-9206-f53f2cacbddb", "createdAt": "2020-10-26T07:58:30.395331+00:00", "updatedAt": "2020-10-26T07:58:30.395331+00:00", "application": "CBEMV", "bankCode": "11899", "rank": null, "contract": "4278155", "merchantLabel": null}, 
        #      {"uuid": "31004942-506a-4313-b51f-0a4ca98035e3", "createdAt": "2020-10-26T07:58:30.397852+00:00", "updatedAt": "2020-10-26T07:58:30.397852+00:00", "application": "CBEMV", "bankCode": "30004", "rank": null, "contract": "4358988", "merchantLabel": null}, 
        #      {"uuid": "aaf82ceb-44a2-4a1e-96ca-cc7ab037d1cc", "createdAt": "2020-10-26T07:58:30.412461+00:00", "updatedAt": "2020-10-26T07:58:30.412461+00:00", "application": "CBEMV", "bankCode": "30004", "rank": null, "contract": "4290394", "merchantLabel": null}, 
        #      {"uuid": "74161e3e-6628-4452-8cd6-2eeb916c71b7", "createdAt": "2020-10-26T07:58:30.417437+00:00", "updatedAt": "2020-10-26T07:58:30.417437+00:00", "application": "CONECS", "bankCode": "10000", "rank": null, "contract": "1207925", "merchantLabel": null}, {"uuid": "05216f24-b58d-434b-8f1b-406030fe1713", "createdAt": "2020-10-26T07:58:30.420862+00:00", "updatedAt": "2020-10-26T07:58:30.420862+00:00", "application": "CBCLESS", "bankCode": "30003", "rank": null, "contract": "3619346", "merchantLabel": null},
        #       {"uuid": "6b1afa41-568e-4cb4-91a8-443b5c0e4652", "createdAt": "2020-10-26T07:58:30.433624+00:00", "updatedAt": "2020-10-26T07:58:30.433624+00:00", "application": "CBEMV", "bankCode": "11899", "rank": null, "contract": "4278622", "merchantLabel": null}], 
        #      "users": [{"uuid": "cec9a8d1-4162-4bb5-a5c7-859340babe63", "email": "thierry.cade@ingenico.com", "createdAt": "2020-10-26T07:58:28.204615+00:00", "updatedAt": "2020-10-26T08:00:23.438885+00:00", "organization": "fb10e478-7981-4cd7-8fb0-d980e7412dc7", "phoneNumber": null, "isAdmin": true}]}
        
    def modifierLibelleContratCom(self,compteUUID,contratUUID,newLabel):
        res = self.callAPI("PUT",f'/contractsAccesses/{compteUUID}/contracts/{contratUUID}',jsonData={'merchantLabel': newLabel} ).json()
        #res = requests.put(self.URI+f'/contractsAccesses/{compteUUID}/contracts/{contratUUID}',json={'merchantLabel': newLabel}, headers=self.HEADERS, cookies=self._cookies)
        #{"success": true}
        assert res['success']
        
    #Attention rank à spécifier sur 3 caractères "001"
    def ajouteUnContratAunCompteAccess(self,compteUUID, listContractRankLabel):
        res = self.callAPI("POST",f'/contractsAccesses/{compteUUID}/contractProfiles',jsonData=listContractRankLabel).json()
        # res = requests.post(self.URI+f'/contractsAccesses/{compteUUID}/contractProfiles',json=listContractRankLabel, headers=self.HEADERS, cookies=self._cookies).json()
        assert res['success']
        #[{"contract":"43673d25-3699-49a3-a989-0fef2c500453","rank":null,"merchantLabel":"testSQ"}]

    def supprimerUnContratDunCompteDAccess(self,compteUUID,contractID):
        res = self.callAPI("DELETE", f'/contractsAccesses/{compteUUID}/contracts/{contractID}').json()
        #res = requests.delete(self.URI+f'/contractsAccesses/{compteUUID}/contracts/{contractID}', headers=self.HEADERS, cookies=self._cookies).json()
        assert res['success']

    def supprimerTousLesContratsDunCompteDaccess(self,compteUUID):
        details = self.detailsCompteDacces(compteUUID)
        if details['contracts']:
            for c in details['contracts']:
                self.supprimerUnContratDunCompteDAccess(compteUUID=details["uuid"],contractID=c["uuid"] )

    def connectAs(self,userId):
        res = self.callAPI("POST",f"/users/{userId}/loginAs")
        URL = res.headers['x-location'] #https://grandfrais.payview.fr/#/login?flt=48cc3955-0acc-4f60-ac13-020f635ebcd9

        m = re.match(r'^https://(?P<ssdomain>.*).payview.fr/#/login\?flt=(?P<token>.*)$', URL) #le ? est un caractère special doit etre échappé: . ^ $ * + ? { } [ ] \ | ( )
        assert m,f"Pas de matching sur URL {URL}"
        self._ssdomainAs = m.group('ssdomain')
        
        self.LoginAs(m.group('token'))

#https://test.payview.fr/api/contractsAccesses/35f06d48-03a6-43d1-8235-5da9b4dd0008/contracts/11701f76-7289-48c7-8e6e-d11f46accb10
    def contratsDisponibles(self,pageSize=50, offset=0,numContrat=None,bankCode=None,rank=None,application=None, partSN = None):
        params = { 'limit':pageSize,'offset':offset,'provider':self._orgaUUID }
        if numContrat:
            params['number']=numContrat
        if bankCode:
            params['bankCode'] = bankCode
        if rank:
            params['rank'] = rank
        if application:
            params['application']=application
        if partSN:
            params['posTerminal.serialNumber'] = partSN
        res = self.callAPI('GET','/merchantContractsFromProvider', query= params ).json()
        # res = requests.get(self.URI+'/merchantContractsFromProvider',params = params, headers=self.HEADERS, cookies=self._cookies)
        # res = res.json()

        if not res['items']:
            return None
        else:
            return res['items']

#  RES:       {
# 	"items": [
# 		{
# 			"uuid": "b8bfbba4-f3de-43c2-92b3-3fdebc69538f",
# 			"application": null,
# 			"number": "6262079",
# 			"rank": "002",
# 			"label": "-",
# 			"bankCode": "30066",
# 			"x25Address": "196358779",
# 			"itp": "193551310711",
# 			"cbVersion": null,
# 			"registrationNumber": "55201420101303",
# 			"legalAndBusinessName": "SELECTA",
# 			"lastSourceIp": "212.243.142.172",
# 			"lastIccid": null,
# 			"lastSslVersion": "TLSv1.2",
# 			"lastConnectionType": "ETHERNET",
# 			"lastRemoteCollectionDate": "2020-10-27T15:59:34.182000+00:00",
# 			"lastConnectionDate": "2020-10-27T15:59:34.182000+00:00",
# 			"createdAt": "2020-10-27T16:00:05.324357+00:00",
# 			"updatedAt": "2020-10-27T16:00:05.324357+00:00",
# 			"posTerminal": {
# 				"uuid": "966d975e-4e42-4875-aae4-ac93df1dcb2f",
# 				"serialNumber": "SE15728903",
# 				"manufacturer": "INGENICO"
# 			},
# 			"posTerminalSubscription": "498eb9c0-8b54-497c-9e34-6d84588a2791"
# 		}	
# 	],
# 	"totalResults": 35
# }

# GET  /merchantContractsFromDirectProvider?provider=fb10e478-7981-4cd7-8fb0-d980e7412dc7&limit=20&offset=0
# Donne les contrats disponibles sur ce client
def exportListeVersExcel(filePath, liste, titre="export"):
    if not liste:
        logging.error(f"Liste vide pour export {titre}")
        return

    tabCli = tablib.Dataset(title=titre, headers= liste[0].keys() )

    for l in liste:
        tabCli.append( l.values() )

    with open( filePath, mode='wb') as f:  #PermissionError si déjà ouvert
            f.write(tabCli.export('xlsx'))

def setup_logging():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    #Console Logger
    consoleHandler = logging.StreamHandler()
    consoleFormatter = colorlog.ColoredFormatter(
        "%(log_color)   s%(message)s",
        datefmt=None,
        reset=True,
        log_colors={
            'DEBUG':    'cyan',
            'INFO':     'green',
            'WARNING':  'yellow',
            'ERROR':    'red',
            'CRITICAL': 'red',
        }
    )
    consoleHandler.setFormatter(consoleFormatter)
    consoleHandler.setLevel(logging.INFO)
    
    if (logger.hasHandlers()):
        logger.handlers.clear()

    logger.addHandler(consoleHandler)
    #File Logger
    dossierTests = os.path.abspath(   os.path.dirname( __file__))
    dossierLogs = os.path.join(dossierTests, 'logs')

    if not os.path.isdir(dossierLogs):
        os.makedirs(dossierLogs)

    logFilePath= os.path.join( dossierLogs, f'{time.strftime("%Y%m%d_%Hh%M")}_tests.log')

    fileHandler =logging.FileHandler(filename=logFilePath, mode='a', encoding="utf-8", delay=False)
    fileHandler.setFormatter( logging.Formatter('%(levelname)s :: %(message)s') )
    fileHandler.setLevel(logging.DEBUG)
    logger.addHandler(fileHandler)
    return f"Logs sous {logFilePath}"
