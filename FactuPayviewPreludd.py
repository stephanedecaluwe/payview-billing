import glob
import locale
import logging
import os
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime
from typing import DefaultDict,List
from dataclasses import dataclass,field,asdict
import tablib
import threading

from APIPayview import APIPayview
from APIFactuPayView import getFactuFilesPayView
from mshAPI import BrowseBillingException, DownloadBillingAPI
from toolsFactu import (exportListesVersExcel, readCsvOrExcel, setLogger, showCallsAndTime)

##################### EMAIL MOT DE PASSE PAYVIEW ADMIN ###########################
PAYVIEW_ADMIN_EMAIL= 'laure.ingenico@gmail.com' #remplacer par l'email admin PayView de Laure
PAYVIEW_ADMIN_MDP='Factu@2021' #Remplacer par le mot de passe admin PayView de Laure

##################### CHEMINS DES FICHIERS D'ENTREES ##############################
CLIENTS_IGNORES                     = r'C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\08.Factu_mai\01.Inputs\clients_non_facturés_Preludd.txt'
FICHIER_SIM_PRET                    = r"C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\08.Factu_mai\01.Inputs\SIMs_de_pret_Preludd.xlsx"
FICHIER_CORRESPONDANCE_NOMS_CLIENTS = r'C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\08.Factu_mai\01.Inputs\correspPayViewPassPort.xlsx'

##################### CHEMINS DES FICHIERS GENERES ##############################
DOSSIER_GENERATION_RESULTATS        = r'C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\08.Factu_mai\03.resultat_factu_Preludd'

######################## VAR GLOBALES #################################
REGEX_FORFAIT_SIM = re.compile(r'^[^0-9]*(?P<taille>\d{1,3})MB$')  #extraire taille de 'Ingenico 2MB' par ex
FORFAITS_SIM_MB_POSSIBLES=[1,2,5,10,50,100]

LISTE_CLIENTS_RIEN_A_FACTURER = []

#Rempli par lecture SIMs de prêt
dicSSNToSimPret = {}

# Nom de client PayView vers infos de facturation
PayViewClientToFactData = {}
# Nom grossiste vers noms des sous clients PayView
GrossistesData= {} #nom vers FactuGrossiste

#Rempli par lecture CLIENTS_IGNORES
NomClientsIgnorésMinuscules = []

#Status TPE et SIM vers statut facturation
PayviewTPEStatusToFactu= {'Activé':True,'En attente':False,'En stock':False,'Résilié':False }
PayViewSIMStatusToFactu={'Activée':True, 'Préactivée':False,"En cours d'activation":False,"Inactive":False,"Résiliée":False,"Suspendue":True}

PassPortSIMStatusToFactu={'Activated':True, "Inventory":False,"Deleted":False,"Suspended":True}

SimsPassPortSansForfaits = {'8934072179002318175':1, '8934072179002318183':1}

columnsOnlyForIngenico = ['nbSimsPassPort','nbSimsO2','nbSimsBouygues','nbTpesPassPort']

_nowCached = f'{datetime.now().strftime("%Y%m%d_%Hh%M")}'
moisFacturationString = None

def getNowStr():
    return _nowCached

@showCallsAndTime
def readFichierSimsPret():
    RegexSsn = re.compile(r"^(?P<SSN>\d{19,20})$",re.IGNORECASE)
    logging.info(f"Lecture fichier sim de prêt {FICHIER_SIM_PRET}") 
    
    for numéroLigneExcel,l in enumerate(readCsvOrExcel(FICHIER_SIM_PRET),start=2):
        ssnLu = str(l['SSN']).strip()	

        m = RegexSsn.match(ssnLu )
        if( m== None ):
            logging.warning(f"Pas de matching SSN {ssnLu} pour SIM Pret à la ligne {numéroLigneExcel}")
            SSN =ssnLu
        else:
            SSN =m.group('SSN')

        if SSN in dicSSNToSimPret:
            logging.warning(f"SIM de pret {SSN} à la ligne {numéroLigneExcel} déjà présente dans fichier de prêt")
            continue

        dicSSNToSimPret[SSN]= {'commentaire': l['Commentaires'], 'clientFichierPret':l['Prêté à (client ou collaborateur.trice)'],'début':l['Date début de prêt'],'fin':l['Date fin de prêt']}
        
    logging.info(f"Fichier de prêt: {len(dicSSNToSimPret)} sims sans doublons")        

@dataclass
class SIM:
    iccid:str
    fournisseur:str
    label:str
    status: str
    forfaitMB:int
    conso_ko:int
    activation_date:str
    surPayView:bool

@dataclass
class TPE:
    SN:str
    label:str
    nbConnexions: int
    reporting: str #'Oui' ou 'Non'
    surPayView:bool

@dataclass
class FactuClient:
    nomPayView: str
    nomPassport: str
    codeClientSap:str
    contratSap:str

    fromGrossiste: str = field(default="")
    forceReporting:bool = field(default=False)
    BU:str = field(default='')
    
    nbReporting: int = field(default=0, init=False)
    TPEs: List[TPE] = field(default_factory=list, init=False)
    NbSIMsParForfait: DefaultDict[int,int] = field( init=False) # 1 -> 25, 2->12
    ConsoParForfait: DefaultDict[int,int] = field( init=False) #Totale conso en Ko par type de forfait en MB
    SIMsPret: List[SIM] = field(default_factory=list, init=False)
    SIMs: List[SIM] = field(default_factory=list, init=False)

    FactuObj: dict  =field(default_factory=dict, init=False)
    FactuDetailsObj: dict  =field(default_factory=dict, init=False)
    #Sera rempli par calcul Billing
    RienAFacturer: bool =field(default=None, init=False)

    def __post_init__(self):
        self.SIMsParForfait = defaultdict(int)
        self.ConsoParForfait = defaultdict(int)
        assert self.nomPayView,"Le nom PayView du client ne peut être vide"

        if self.fromGrossiste:
            nomGrossiste  =self.fromGrossiste
            if nomGrossiste not in GrossistesData:
                GrossistesData[nomGrossiste] = FactuGrossiste(nomGrossiste)

            GrossistesData[nomGrossiste].ajouteSousClient(self)
            self.contratSap = self.codeClientSap = ""

    def calculeBilling(self):
        if self.FactuObj:
            return logging.warning(f"factuClient calculeBilling déjà fait sur {self.nomPayView}")

        self.FactuObj['Client']             = self.nomPayView
        self.FactuObj['Grossiste']          = self.fromGrossiste
        self.FactuObj['BU']                 = self.BU
        self.FactuObj['CODE CLIENT SAP']    = self.codeClientSap
        self.FactuObj['CONTRAT SAP']        = self.contratSap
        self.FactuObj['PAS_IP500']          = 0 #pour mettre le champ en 3e, sera rempli plus tard
        self.FactuObj['PAS_REPORTING']      = self.nbReporting

        self.FactuDetailsObj['Client']      = self.nomPayView
        self.FactuDetailsObj['Grossiste']   = self.fromGrossiste

        surconsoSimsCeClientKo = 0
        for tailleForfaitMB in FORFAITS_SIM_MB_POSSIBLES:
            codeSAP = f"PAS_SIM{tailleForfaitMB}_{500 if tailleForfaitMB<5 else 0}_N"
            nbSimsActivesCeForfait = self.SIMsParForfait[tailleForfaitMB]
            self.FactuObj[codeSAP] = nbSimsActivesCeForfait

            poolMB = nbSimsActivesCeForfait * tailleForfaitMB #en MB
            surconsoKo =max(0, int(self.ConsoParForfait[tailleForfaitMB] -1024*poolMB))

            self.FactuDetailsObj[f'Pool_{tailleForfaitMB}_MB'] =poolMB
            self.FactuDetailsObj[f'Conso_{tailleForfaitMB}_MB'] =int(self.ConsoParForfait[tailleForfaitMB]/1024)

            surconsoSimsCeClientKo  += surconsoKo
            
        nbSimsActivesTotalCeClient  = len(self.SIMsPret) + len(self.SIMs)
        nbTpesActifs =  len(self.TPEs)
        nbTpesFactures = max(0,nbTpesActifs- nbSimsActivesTotalCeClient)
        
        self.FactuDetailsObj['TPEs_actifs']     = nbTpesActifs
        self.FactuObj['PAS_IP500']              = nbTpesFactures
        self.FactuDetailsObj['TPEs_factures']   = nbTpesFactures
        
        self.FactuObj['PAS_SIM_OVERFEE'] = surconsoSimsCeClientKo
        #Ajout des détails Factu finaux
        self.FactuDetailsObj['nbSimsActives']   = nbSimsActivesTotalCeClient
        self.FactuDetailsObj['dontNbSimsDePret'] = len(self.SIMsPret)

        self.FactuDetailsObj['nbSimsPassPort']  = len([s for s in self.SIMs if not s.surPayView]) + len([s for s in self.SIMsPret if not s.surPayView])
        self.FactuDetailsObj['nbSimsO2']        = len([s for s in self.SIMs if s.fournisseur=='O2'])
        self.FactuDetailsObj['nbSimsBouygues']  = len([s for s in self.SIMs if s.fournisseur=='bouyguesTelecom'])

        self.FactuDetailsObj['nbTpesPassPort']  = len([t for t in self.TPEs if not t.surPayView])

        self.RienAFacturer = not any( (k.startswith('PAS') and val >0 for k,val in self.FactuObj.items() ) )
        if self.RienAFacturer:
            logging.warning(f"Rien à facturer client '{self.nomPayView}'")
            LISTE_CLIENTS_RIEN_A_FACTURER.append(self.nomPayView)
        
    def ExportDetailsExcel(self,dossier):
        listeDataTitre = []
        if self.RienAFacturer == None:
            self.calculeBilling()

        if self.RienAFacturer:
            return
            
        listeDataTitre.append(([self.FactuObj] , 'Facturation'))
        listeDataTitre.append(([self.FactuDetailsObj] , 'détailsPoolsSims'))

        if self.TPEs:
            listeDataTitre.append(([asdict(t) for t in self.TPEs] , 'TPEs actifs'))
        if self.SIMs:
            listeDataTitre.append(([asdict(s) for s in self.SIMs] , 'SIMs actives'))
        if self.SIMsPret:
            listeDataTitre.append(([asdict(s) for s in self.SIMsPret] , 'SIMs prêtées'))

        if self.fromGrossiste:
            dossier = os.path.join(dossier,self.fromGrossiste )
            os.makedirs(dossier, mode=0o777,exist_ok=True)

        fPath = os.path.join(dossier,f"détails_{moisFacturationString}_{self.nomPayView}.xlsx")
        exportListesVersExcel(fPath, listeDataTitre, colNamesTodel=columnsOnlyForIngenico)

    def ajouteTPE(self,tpe):
        if self.forceReporting:
            tpe.reporting = 'Oui'
        
        self.TPEs.append(tpe)
        if tpe.reporting.lower() == 'oui':
            self.nbReporting +=1

    def ajouteSIM(self,sim:SIM):
        if sim.iccid in dicSSNToSimPret:
            logging.warning(f"Sim de prêt {sim.iccid} chez {self.nomPayView}")
            self.SIMsPret.append(sim)
        else:
            self.SIMs.append(sim)
            self.SIMsParForfait[sim.forfaitMB] += 1
            self.ConsoParForfait[sim.forfaitMB] += sim.conso_ko

@dataclass
class FactuGrossiste:
    nom:str
    
    BU:str= field(default='',init=False)
    codeSAP:str = field(default="",init=False)
    contratSAP:str = field(default="",init=False)
    FactuSsClients: list = field(default_factory=list, init=False)
    #Utilisé que pour savoir s'il faut ajouter le sous client découverte par API
    SousClients : list = field(default_factory=list, init=False)
    # Privés:
    FactuGlobale: FactuClient = field(default=None, init=False)
    ListeFactuGlobale: list = field(default_factory=list, init=False)
    ListeDetailsGlobale: list = field(default_factory=list, init=False)

    def getFactuGlobale(self):
        if not self.FactuGlobale:
            self.FactuGlobale = FactuClient(self.nom, "",self.codeSAP,self.contratSAP,fromGrossiste="",BU=self.BU)
            self.ListeFactuGlobale.append(self.FactuGlobale.FactuObj)
            self.ListeDetailsGlobale.append(self.FactuGlobale.FactuDetailsObj)

            for factSsClient in self.FactuSsClients:
                self.FactuGlobale.SIMs += factSsClient.SIMs
                self.FactuGlobale.TPEs += factSsClient.TPEs
                self.FactuGlobale.SIMsPret += factSsClient.SIMsPret
                self.FactuGlobale.nbReporting += factSsClient.nbReporting
                
                for k,v in factSsClient.SIMsParForfait.items():
                    self.FactuGlobale.SIMsParForfait[k] += v 
                    self.FactuGlobale.ConsoParForfait[k] += factSsClient.ConsoParForfait[k]

                self.ListeFactuGlobale.append(factSsClient.FactuObj)
                self.ListeDetailsGlobale.append(factSsClient.FactuDetailsObj)
            
            self.FactuGlobale.calculeBilling()

        return self.ListeFactuGlobale, self.ListeDetailsGlobale
        
    def makeExcelGlobalGrossiste(self, dossierDetails):
        self.getFactuGlobale()
        dossierCeGrossiste = os.path.join(dossierDetails, self.nom)
        fichierBilanGr = os.path.join(dossierCeGrossiste, f'{moisFacturationString}_BILAN_grossiste_{self.nom}.xlsx')
        exportListesVersExcel(fichierBilanGr, [(self.ListeFactuGlobale,'Facturation'),(self.ListeDetailsGlobale,'Détails')],colNamesTodel=columnsOnlyForIngenico)
        #Crée zip pour chaque dossier grossiste
        shutil.make_archive(os.path.join(dossierDetails,moisFacturationString+"grossiste_"+self.nom ),"zip", dossierCeGrossiste)

    def IsClientNameInSousClient(self, clientName):
        return any( (c.nomPayView == clientName for c in self.FactuSsClients) )

    def ajouteSousClient(self,sousClient:FactuClient):
        self.SousClients.append(sousClient.nomPayView)
        self.FactuSsClients.append(sousClient)

        if not self.codeSAP and sousClient.codeClientSap:
            self.codeSAP = sousClient.codeClientSap
        if not self.contratSAP and sousClient.contratSap:
            self.contratSAP = sousClient.contratSap
        if not self.BU and sousClient.BU:
            self.BU = sousClient.BU
        if not sousClient.BU:
            sousClient.BU = self.BU

def demandeAttention(errorMsg:str):
    logging.error(errorMsg)
    input("ATTENTION ??")

# @showCallsAndTime
def readCorrespondanceNomClients():
    global PayViewClientToFactData,NomClientsIgnorésMinuscules

    if PayViewClientToFactData:
        return

    for l in readCsvOrExcel(FICHIER_CORRESPONDANCE_NOMS_CLIENTS):
        nomPayView      = l['clientName']
        nomPassport     = l['NomPassPort'].strip() if l['NomPassPort'] else ""

        if nomPayView in PayViewClientToFactData:
            demandeAttention(f"{nomPayView} déjà présent dans liste clients PayView")

        PayViewClientToFactData[nomPayView] = FactuClient(nomPayView,nomPassport,l['CODE CLIENT SAP'],l['CONTRAT SAP'], fromGrossiste=l['Grossiste'] , forceReporting=l['ForceFactuReporting'],BU=l['BU'] )
        
    # #Lire tous les clients par l'API de PayView
    api =APIPayview("admin",PAYVIEW_ADMIN_EMAIL,PAYVIEW_ADMIN_MDP)
    api.Login()
    res = api.getAllClients()
    for c in res:
        name = c['legalName']
        logging.info(f"Vérification présence client PayView {name}")

        if not c['grossiste']:
            if name not in PayViewClientToFactData:
                if input(f'ajouter `{name}` à la liste de clients PayView O/N ?').lower() =='o':
                    PayViewClientToFactData[name]= FactuClient(name,nomPassport="",codeClientSap='?',contratSap='?')
        else:
            startGrName = c['grossiste'].split()[0]
            for gr in GrossistesData.keys():
                if gr.startswith(startGrName):
                    if not GrossistesData[gr].IsClientNameInSousClient(name):
                        if input(f"Ajouter `{name}` au grossiste {gr} O/N ?").lower().strip() == 'o':
                            PayViewClientToFactData[name]= FactuClient(name,nomPassport="",codeClientSap='?',contratSap='?',fromGrossiste=gr)
                    break
            else:
                logging.error(f"{startGrName} non trouvé dans les grossistes")
                sys.exit()

    #Clients ignorés
    with open(CLIENTS_IGNORES,'r',encoding='utf8' ) as f:
        NomClientsIgnorésMinuscules = [c.strip().lower() for c in f.readlines()]
   
    logging.info(f"correspondance noms clients {len(PayViewClientToFactData)} clients trouvés - Nb clients à ignorer: {len(NomClientsIgnorésMinuscules)}") 

def lectureDossier(isPayView:bool, directory:str):
    RegexTpeFile = re.compile(r"^(?P<name>.*)_terminal-connections-for-(?P<ID>.*).csv$",re.IGNORECASE )
    RegexSimFile = re.compile(r"^(?P<name>.*)_sim-status-for-(?P<ID>.*).csv$",re.IGNORECASE )

    def getPayViewNameFromPassPortName(passportClient):
        """
        Recherche le nom PayView à partir du nom PassPort du client
        """
        for k,v in PayViewClientToFactData.items():
            if v.nomPassport == passportClient:
                return k
        
        demandeAttention(f"client PassPort '{passportClient}' non trouvé dans clients PayView")
        return None

    def _lectureFichierTpe(isPayView:bool, filePath:str, fileName:str):
        TradPassToPay = {'Number Of Connections':'number_of_connections','Serial Number':'serial_number','Pos Name':'tpe_label','Customer Name':'fournisseur', 'Reporting Service Access':'reporting_service_access' }

        if not isPayView:
            mtpe = RegexTpeFile.match(filename)
            if not mtpe:
                raise ValueError(f"Nom de fichier PassPort {fileName} au mauvais format TPE file")
            clientPassPort = mtpe.group('name')
            if clientPassPort.strip().lower() in NomClientsIgnorésMinuscules:
                return logging.warning(f"Fichier TPE {fileName} client ignoré")
            clientPayView = getPayViewNameFromPassPortName(clientPassPort)
            if not clientPayView:
                return demandeAttention(f"client {clientPassPort} de PassPort pas dans liste correspondance")

        for t in readCsvOrExcel(filePath,forceEncoding='utf_16' if isPayView else None):
            if isPayView and not PayviewTPEStatusToFactu[t['subscription_status']]:
                continue

            if not isPayView:
                for k,v in TradPassToPay.items():
                    t[v] = t[k]

            nbConnexions = int(t['number_of_connections'])
            if nbConnexions <1:
                continue #non facturé car pas de connexions

            if isPayView:
                clientPayView = t['fournisseur']

            if clientPayView not in PayViewClientToFactData:
                demandeAttention( f"lecture TPE {filePath} client {clientPayView} inconnu ??" )
                continue

            PayViewClientToFactData[clientPayView].ajouteTPE(TPE(t['serial_number'],t['tpe_label'],nbConnexions,t['reporting_service_access'],isPayView ))

    def _lectureFichierSim(isPayView:bool, filePath:str,fileName:str):
        TradPassToPay = {'Status':'status','Name':'sim_label','Operator':'operator','SSN':'iccid', 'Contract':'forfait','Activation Date':'activation_date'}
        inconnus = set()

        if not isPayView:
            msim = RegexSimFile.match(filename)
            if not msim:
                raise ValueError(f"Nom de fichier PassPort {fileName} au mauvais format SIM file")
            clientPassPort = msim.group('name')
            if clientPassPort.strip().lower() in NomClientsIgnorésMinuscules:
                return logging.warning(f"Fichier SIM {fileName} client ignoré")
            clientPayView = getPayViewNameFromPassPortName(clientPassPort)
            if not clientPayView:
                return demandeAttention(f"client {clientPassPort} de PassPort pas dans liste correspondance")

        #lecture du fichier
        for l in readCsvOrExcel(filePath,forceEncoding='utf_16' if isPayView else None):
            if not isPayView:
                for k,v in TradPassToPay.items():
                    l[v] = l[k]
                l['sim_data_volume'] = int(l['Sim Volume To Terminal']) + int(l['Sim Volume From Terminal'])

            status = l['status']

            if  (isPayView and not PayViewSIMStatusToFactu[status]) or (not isPayView and not PassPortSIMStatusToFactu[status]):
                continue   #statut SIM non facturé 

            if isPayView:    
                clientPayView = l.get('sim_final_client',None) or l['fournisseur'] 
                    
                if clientPayView not in PayViewClientToFactData:
                    nomClientIngenico = l['client_name']
                    if nomClientIngenico in PayViewClientToFactData: #Ce n'est pas un client d'un grossiste
                        clientPayView = nomClientIngenico
                    else: #C'est un client d'un grossiste
                        if clientPayView not in inconnus:
                            inconnus.add(clientPayView)
                            demandeAttention(f"Nom du client '{clientPayView}' pour fichier sim inconnu '{fichierSim}' on ignore la ligne")
                            continue
            else:
                if l['operator']== 'sierraWireless':
                    continue #Les sims Sierra ne sont lues que dans PayView pas Passport
            iccid = l['iccid']
            nomForfait=l['forfait']
            tailleForfaitSIM_MB = 0

            resF = REGEX_FORFAIT_SIM.match(nomForfait)
            if not resF:
                tailleForfaitSIM_MB = {'1024':1,'2048':2,'5120':5}.get(nomForfait,None)

                if not tailleForfaitSIM_MB:
                    if iccid in SimsPassPortSansForfaits:
                        tailleForfaitSIM_MB = SimsPassPortSansForfaits[iccid]
                    else:
                        raise ValueError(f"forfait SIM inconnu '{nomForfait}' dans {filePath}")
            else:
                tailleForfaitSIM_MB = int(resF.group('taille'))
            
            assert tailleForfaitSIM_MB in FORFAITS_SIM_MB_POSSIBLES, f"Mauvaise taille forfait SIM {tailleForfaitSIM_MB}"

            consoKo = int( int(l['sim_data_volume'])/1024)

            PayViewClientToFactData[clientPayView].ajouteSIM(SIM(iccid,l['operator'],l['sim_label'],l['status'],tailleForfaitSIM_MB,consoKo,l['activation_date'],isPayView ))

    if not directory:
        return logging.info(f'Pas de dossier {"PayView" if isPayView else "PassPort"} à lire')
    
    logging.info(f'Lecture dossier {"PayView" if isPayView else "PassPort"}')
    for f in glob.glob(os.path.join(directory,'*.csv')):
        filename = os.path.basename(f)
        if 'sim-status' in filename:
            _lectureFichierSim(isPayView, f,filename )
        elif (not isPayView and 'terminal-connections' in filename) or (isPayView and 'poi-connections' in filename):
            _lectureFichierTpe(isPayView, f,filename )

def readData(payViewDirectory, passPortZipDirectory):
    if not NomClientsIgnorésMinuscules:
        readCorrespondanceNomClients()
    
    if not dicSSNToSimPret:
        readFichierSimsPret()

    lectureDossier(isPayView=False, directory=passPortZipDirectory)
    lectureDossier(isPayView=True,  directory=payViewDirectory)

@showCallsAndTime
def génèreFacturationComplete(dossierDataPassPort,dossierDataPayView,moisString, dossierGénéCetteFactu, inclusNonFacturés=False):
    #Lecture données
    readData(dossierDataPayView,dossierDataPassPort)

    listeToExcelFactu = defaultdict(list)
    listeToExcelDetails = defaultdict(list)

    dossierDétails = os.path.join(dossierGénéCetteFactu,"Détails")
    os.makedirs(dossierDétails, mode=0o777, exist_ok=True)

    #Traitement de chaque client FINAL
    for clientName,v in PayViewClientToFactData.items():
        logging.info(f"Traitement factu {clientName}")
        v.ExportDetailsExcel(dossierDétails)

        qqchoseAFacturer = not v.RienAFacturer
        duGrossiste = v.fromGrossiste

        if not duGrossiste and (qqchoseAFacturer or (not qqchoseAFacturer and inclusNonFacturés)):
            listeToExcelFactu[v.BU].append(v.FactuObj)
            listeToExcelDetails[v.BU].append(v.FactuDetailsObj)
    
    #Traitement de chaque grossiste et ses sous clients
    for nomGrossiste,factuGrossiste in GrossistesData.items():
        logging.info(f"Traitement factu grossiste {nomGrossiste}")
        lignesFactu,lignesDetails = factuGrossiste.getFactuGlobale()
        listeToExcelFactu[v.BU].append(lignesFactu)
        listeToExcelDetails[v.BU].append(lignesDetails)

        factuGrossiste.makeExcelGlobalGrossiste(dossierDétails)

    #Generation fichier factu globale
    books = tablib.Databook()

    for (liste,titre) in [(listeToExcelFactu,'Facturation'),(listeToExcelDetails,'Détails')]:
        ds = tablib.Dataset( title=titre,headers=list(liste.values())[0][0].keys() )

        for _,v in liste.items():
            for l in v:
                if isinstance(l,list):
                    ds.extend( [li.values() for li in l] )
                else:
                    ds.append( l.values() )

        books.add_sheet( ds )

    #Ecriture du fichier Excel
    globalFactuFile = None
    if books.size: #Vide
        globalFactuFile = os.path.join(dossierGénéCetteFactu,f'{getNowStr()}_{moisFacturationString}_facturationGlobale.xlsx' )
        with open( globalFactuFile, mode='wb') as f:
                f.write(books.export('xlsx'))

    #Crée zip pour ADV
    shutil.make_archive(os.path.join(DOSSIER_GENERATION_RESULTATS,f"{getNowStr()}_factuPayView_{moisFacturationString}_pourADV"),"zip", dossierGénéCetteFactu)
    return globalFactuFile

@showCallsAndTime
def génèreFactu(année:int,mois:int):
    global moisFacturationString
    moisFacturationString =  datetime(année,mois,1).strftime('%B_%Y') # 'octobre_2020'
    nomDossier = f'{getNowStr()}_facturation_{moisFacturationString}'

    dossierGénéCetteFactu = os.path.join(DOSSIER_GENERATION_RESULTATS, nomDossier )
    os.makedirs(dossierGénéCetteFactu, mode=0o777, exist_ok=False)

    readCorrespondanceNomClients()

    # Remplacer les valeurs par None sur les 2 lignes pour récupérer les données de PassPort et PayView
    #PASSPORT_ZIP_DIRECTORY      =   r"C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\DataPassPort"
    #PAYVIEW_DATA_DIRECTORY      =   r"C:\Users\lbroegg\Ingenico_Workspace\202011_Facturation\DataPayView"
    PASSPORT_ZIP_DIRECTORY      =   None
    PAYVIEW_DATA_DIRECTORY      =   None

 
    if not PAYVIEW_DATA_DIRECTORY: #not ALL_SIMS_PAYVIEW or not ALL_TPES_PAYVIEW:
        PAYVIEW_DATA_DIRECTORY = os.path.join(dossierGénéCetteFactu,"DataPayView")
        os.makedirs(PAYVIEW_DATA_DIRECTORY, mode=0o777, exist_ok=False)
        getFactuFilesPayView(PAYVIEW_DATA_DIRECTORY, mois, année, ignoredListMinuscule=NomClientsIgnorésMinuscules )

    #Génère la facturation complète
    globalXlFile = génèreFacturationComplete(dossierDataPassPort =PASSPORT_ZIP_DIRECTORY ,dossierDataPayView=PAYVIEW_DATA_DIRECTORY,moisString=moisFacturationString, dossierGénéCetteFactu=dossierGénéCetteFactu, inclusNonFacturés=False)

    if LISTE_CLIENTS_RIEN_A_FACTURER:
        logging.info("Ecriture fichier des clients sans rien à facturer")
        with open(os.path.join(dossierGénéCetteFactu, 'clientsRienAFacturer.txt'),'w',encoding='utf8') as f: f.write("\n".join(LISTE_CLIENTS_RIEN_A_FACTURER))

    if globalXlFile:     #Ouverture de l'Excel
        def ouvreExcel():
            os.system(globalXlFile)

        t=threading.Thread(target =ouvreExcel)
        t.start()
        t.join(0)

#Début du script
if __name__ == "__main__":
    locale.setlocale(locale.LC_ALL, '') #set local en Français
    setLogger()
    mois =5
    année = 2021
    génèreFactu(année,mois)