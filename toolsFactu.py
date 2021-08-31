import colorlog
import logging
import time
import os
import functools
import tablib
import copy

def showCallsAndTime(func):
    @functools.wraps(func)
    def inner(*args,**kwargs):
        start = time.time()
        
        filenameOnly =""
        if 'filePath' in kwargs:
            filenameOnly = '(' + os.path.basename(kwargs['filePath']) + ')'

        logging.info(f"{func.__name__} DEBUT {filenameOnly}")
        ret = func(*args,**kwargs)
        dur = time.time() - start
        logging.info(f"{func.__name__} FIN ({dur:.2f} secondes)")
        return ret
    return inner

def exportListesVersExcel(filePath, listeDeDataTitreParOnglet, colNamesTodel=None):
    if not listeDeDataTitreParOnglet:
        logging.error(f"Liste vide pour export {filePath}")
        return

    books = tablib.Databook()
    #Population du fichier à écrire
    for liste,titre in listeDeDataTitreParOnglet:
        if len(liste)<1:
            continue
        ds = tablib.Dataset(*[l.values() for l in liste], title=titre, headers= liste[0].keys() )
        if colNamesTodel:
            for col in colNamesTodel:
                if col in ds:
                    del ds[col]

        books.add_sheet( ds )

    #Ecriture du fichier Excel
    if not books.size: #Vide
        return

    with open( filePath, mode='wb') as f:
            f.write(books.export('xlsx'))

def readCsvOrExcel(filePath, forceEncoding=None):
    if filePath.lower().endswith('csv'):
        #Les fichiers PayView: utf-16 et delimiteur tabulation
        #Ficheirs MSH: utf-8 et ;
        encoding = forceEncoding if forceEncoding else 'utf-8'
        delimiter = '\t' if forceEncoding else ';'

        with open(filePath,'r',encoding=encoding) as fh:
            data =tablib.Dataset().load(fh,'csv',delimiter=delimiter )

        if forceEncoding: #Nettoyage des fichiers de Preludd:
            res = copy.copy(data.dict)
            for ligne in res:
                for k,v in ligne.items():
                    if v and type(v)==str:
                        ligne[k] =v.strip('"=')
            return res
    else:
        with open(filePath,'rb' ) as fh:
            data =tablib.Dataset().load(fh,'xlsx' )

    return data.dict

def setLogger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    consoleHandler = logging.StreamHandler() #Console Logger
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
    logDir = os.path.join( os.getcwd(),'logs')
    os.makedirs(logDir,exist_ok=True)
    logFilePath= os.path.join(logDir, f'{time.strftime("%Y%m%d_%Hh%M")}_billing.log')

    fileHandler =logging.FileHandler(filename=logFilePath, mode='a', encoding="utf-8", delay=False)
    fileHandler.setFormatter( logging.Formatter('%(levelname)s :: %(message)s') )
    fileHandler.setLevel(logging.DEBUG)
    logger.addHandler(fileHandler)
