import subprocess
import requests
import json
import time
import os
import pandas as pd
import numpy as np
import ast
import pyodbc
def connessioneDataBase():

     # Configurazione connessione
    server="ntsql013"
    database = "ES_R_42"
    driver = '{ODBC Driver 18 for SQL Server}'  # Verifica se questo driver è installato
    connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};Trusted_Connection=yes;Encrypt=no'

    try:
        conn = pyodbc.connect(connection_string)
        return conn
    except pyodbc.Error as e:
        print(f"Errore di connessione: {e}")
        return None

 
token="hQAMDIovkn4hVK8Lgb/2mDfgoKeNx1pCeawdZh4t4EmXBnMx0yw5klKJVjMSQaGZlfZbr11dNWQxkaWb91GJPEXxMTg5TqWum6lnsZ+qaRDf2Yl/wumY3w=="
idDocumento=''
parametriDaRimuovereDocLeggi=[ "Dati.IdDocumento","Dati.NumRegistraz","Dati.NumDocOriginale"]
parametriDaRimuovereDocCrea=[ "IdDocumento","Documento.IdDocumento","Documento.NumRegistraz","Documento.NumDocOriginale"]
direttivaBase="c:\\Users\\fiore\\Documents\\TEST NON REGRESSIONE\\TEST\\ORV\\"
PercorsoFile=''
nomeCartellaTest="TEST"
cartellaBenchmark=direttivaBase+"BENCHMARK"
isDocumento=False
direttivaWinMerge="C:\\Program Files\\WinMerge\\WinMergeU.exe"
isCreaORVdaOPV=False
isCreaOPV=False



def creaToken():
    url_token = "http://toinsxa0040:9050/auth/App-Sistemi/Token"
    headers = {
        "Content-Type": "application/json"
    }
    payload={
  "User": "CRM",
  "Key": "130f70ae5b44ed4c645aad5cf7cf18ad2413c48e754dbfdec86cc781a26a5624"
 }
    response = requests.post(url_token, headers=headers, data=json.dumps(payload))
    if response.status_code == 200:
        tokenCreato=response.text
        print(f"Token creato: {tokenCreato}")
        return tokenCreato
    else:
     print(f"Errore nella creazione del token: {response.status_code}, {response.text}")
     return None

def checkToken():
    url_checkToken = "http://toinsxa0040:9050/auth/App-Sistemi/CheckToken"
    headers = {
        "Content-Type": "application/json"
    }
    payload={
  "Token": token
    }
    response = requests.post(url_checkToken, headers=headers, data=json.dumps(payload))
    if response.status_code == 200:
        tokenValido = response.json().get("IsValid") 
        if(tokenValido) :
         print(f"Token attivo")
         return tokenValido
        else:
           print(f"Token non attivo")
           return tokenValido
    else:
        print(f"Errore nella creazione dell'ordine: {response.status_code}, {response.text}")
        return None

    

def eseguiQuery(conn,order_id,DBGruppo):
    cursor = conn.cursor()  # Crea un cursore
    query = """
    SELECT StatoDocumento
    FROM DocElencoGen
    WHERE dbgruppo = ? AND IdDocumento = ?
    """
    try:
        # Crea il cursore ed esegui la query
        cursor = conn.cursor()
        cursor.execute(query,(DBGruppo, order_id))
        
        # Recupera il risultato
        risultato = cursor.fetchone()  # Restituisce la prima riga del risultato
        if risultato:
            stato_documento = risultato[0]  # StatoDocumento è il primo elemento della riga
            return stato_documento
        else:
            print("Nessun risultato trovato.")
            return None
    except Exception as e:
        print(f"Errore durante l'esecuzione della query: {e}")
        return None
    



def crea_cartella_con_contatore(direttivaBase, nomeCartella):
    
    if not os.path.exists(direttivaBase):
        os.makedirs(nomeCartella)  # Crea la directory di base se non esiste

    # Trova il contatore massimo tra le cartelle esistenti
    cartelle_esistenti = [
        d for d in os.listdir(direttivaBase)
        if os.path.isdir(os.path.join(direttivaBase, d)) and d.startswith(nomeCartella)
    ]

    contatore = 1
    for cartella in cartelle_esistenti:
        try:
            numero = int(cartella.split("_")[-1])
            contatore = max(contatore, numero + 1)
        except ValueError:
            continue

    # Crea la nuova cartella
    nuova_cartella = os.path.join(direttivaBase, f"{nomeCartella}_{contatore}")
    os.makedirs(nuova_cartella)
    return nuova_cartella

def creaNuovaCartella(direttivaBase, cartellaBase):  
    
    if not os.path.exists(direttivaBase):
        os.makedirs(cartellaBase)  # Crea la directory di base se non esiste
   
    nuova_cartella = os.path.join(direttivaBase, cartellaBase)

    if not os.path.exists(nuova_cartella):
        os.makedirs(nuova_cartella)
        return nuova_cartella
    # Trova il contatore massimo tra le cartelle esistenti
    cartelle_esistenti = [
        d for d in os.listdir(direttivaBase)
        if os.path.isdir(os.path.join(direttivaBase, d)) and d.startswith(cartellaBase)
    ]

    contatore = 1
    for cartella in cartelle_esistenti:
        try:
            numero = int(cartella.split("_")[-1])
            contatore = max(contatore, numero + 1)
        except ValueError:
            continue

    # Crea la nuova cartella
    nuova_cartella = os.path.join(direttivaBase+"\\", f"{cartellaBase}_{contatore}")
    os.makedirs(nuova_cartella)
    return nuova_cartella



def salvaSuFile(NomeFile,directory,risposta):
    if not os.path.exists(directory):
     os.makedirs(directory)

    # Crea il percorso completo del file
    '''split=url.split("/")
    nomeFile=split[-1]+""+split[1]+".txt"'''
    percorso_completo = os.path.join(directory, NomeFile+".txt")
    

    # Apri il file in modalità scrittura e scrivi del testo
    with open(percorso_completo, 'w') as file:
       file.write(json.dumps(risposta, indent=4))


def crea(token,DBgruppo,url,payload,directory,NomeFile,conn):
    url_crea = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }
    
    
    response = requests.post(url_crea, headers=headers, data=json.dumps(payload))
    json_data=response.json()
    
    
    if response.status_code == 200:
        if isDocumento:    
            idDocumento = response.json().get("IdDocumento")  
            print(f"Ordine creato con ID: {idDocumento}")
            json_data=rimuovi_attributi(json_data,parametriDaRimuovereDocCrea)
            json_data=rimuovi_null(json_data)
            salvaSuFile(NomeFile,directory,json_data)
            statoDocumento=-1
            esito=response.json().get("Risposta", {}).get("Esito")
            if(esito==1):
                statoDocumento=-1
                while statoDocumento!=0:
                 statoDocumento=eseguiQuery(conn,idDocumento,DBgruppo)
            return idDocumento
        else:
            json_data=rimuovi_null(json_data)
            salvaSuFile(NomeFile,directory,json_data)

    else:
        print(f"Errore nella creazione dell'ordine: {response.status_code}, {response.text}")
        salvaSuFile(NomeFile,directory,json_data)
        return None
    

def leggi(token,DBgruppo,url,payload,valoreId,directory,NomeFile):
    url_leggi = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }
    if isDocumento:
        if "IdDocumento" in payload:
            payload["IdDocumento"] =valoreId 
            print(f"Modificata la chiave  'IdDocumento' con il valore '{valoreId}'.")
            response = requests.post(url_leggi, headers=headers, data=json.dumps(payload))
        else:
            print(f"La chiave 'IdDocumento' non esiste nel JSON.")
            response = requests.post(url_leggi, headers=headers, data=json.dumps(payload))
    else :
        response = requests.post(url_leggi, headers=headers, data=json.dumps(payload))
    
    

    if response.status_code == 200:
       json_data=response.json()
       if isDocumento:
         json_data=rimuovi_attributi(json_data,parametriDaRimuovereDocLeggi)
       json_data=rimuovi_null(json_data)
       salvaSuFile(NomeFile,directory,json_data)
       
        
    else:
        print(f"Errore nella creazione dell'ordine: {response.status_code}, {response.text}")
        salvaSuFile(NomeFile,directory,response.json())
        return None
    
def aggiorna(token,DBgruppo,url,payload,valoreId,directory,NomeFile,conn):
    url_elimina = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }

    if "Parametri" in payload and "IdDocumento" in payload["Parametri"]:
            payload["Parametri"]["IdDocumento"] =valoreId 
            print(f"Modificata la chiave 'IdDocumento' con il valore '{valoreId}'.")
    else:
        print(f"La chiave 'IdDocumento' non esiste nel JSON.")
    
    response = requests.post(url_elimina, headers=headers, data=json.dumps(payload))
    json_data=response.json()
    if response.status_code == 200:
        if isDocumento:    
            json_data=rimuovi_attributi(json_data,parametriDaRimuovereDocCrea)
            json_data=rimuovi_null(json_data)
            salvaSuFile(NomeFile,directory,json_data)
            statoDocumento=-1
            esito=response.json().get("Risposta", {}).get("Esito")
            print(esito)
            if(esito==1):
                statoDocumento=-1
                while statoDocumento!=0:
                 statoDocumento=eseguiQuery(conn,valoreId,DBgruppo)
                print(statoDocumento)
        else:
            json_data=rimuovi_null(json_data)
            salvaSuFile(NomeFile,directory,json_data)

    else:
        print(f"Errore nell'aggiornamento dell'ordine: {response.status_code}, {response.text}")
        salvaSuFile(NomeFile,directory,json_data)
        return None
 

def elimina(token,DBgruppo,url,payload,valoreId,directory,NomeFile):
    url_elimina = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }
    if isDocumento:
        if "IdDocumento" in payload:
                payload["IdDocumento"] =valoreId 
                print(f"Modificata la chiave 'IdDocumento' con il valore '{valoreId}'.")
        else:
            print(f"La chiave 'IdDocumento' non esiste nel JSON.")
    
    response = requests.post(url_elimina, headers=headers, data=json.dumps(payload))

    if response.status_code == 200:
        print(f"Ordine eliminato con successo")
        
    else:
        print(f"Errore nell'eliminazione dell'ordine: {response.status_code}, {response.text}")
        
    
    salvaSuFile(NomeFile,directory,response.json())


def rimuovi_attributi(json_data,parametriDaRimuovereDoc):
    chiave_esiste=True
    for percorso in parametriDaRimuovereDoc:
        attributi = percorso.split('.')  # Dividi il percorso in chiavi
        temp_data = json_data

        for attributo in attributi[:-1]:  # Naviga fino al penultimo elemento del percorso
            if isinstance(temp_data, dict) and attributo in temp_data:
                temp_data = temp_data[attributo]
            elif isinstance(temp_data, list) and attributo.isdigit() and 0 <= int(attributo) < len(temp_data):
                temp_data = temp_data[int(attributo)]
            else:
                chiave_esiste = False  # Il percorso non esiste
                print(f"Percorso non trovato: {'.'.join(attributi)}")
                break

        
        # Rimuovi l'ultima chiave
        if chiave_esiste:
            if isinstance(temp_data, dict)  and attributi[-1] in temp_data:
             # Rimuovi l'ultima chiave solo se esiste
             temp_data.pop(attributi[-1])
            elif isinstance(temp_data, list):  # Se l'elemento è una lista
                for item in temp_data:
                    if isinstance(item, dict) and attributi[-1] in item:
                        item.pop(attributi[-1])  # Rimuovi la chiave da ogni dizionario nella lista
            else:
                print(f"Chiave '{attributi[-1]}' non trovata")

    return json_data
        
def rimuovi_null(json_data):
    if isinstance(json_data, dict):
        # Itera sui dizionari
        return {
            k: rimuovi_null(v)
            for k, v in json_data.items()
            if v not in (None, 0, "", 0.0, [], {}) and rimuovi_null(v) != {}
        }
    elif isinstance(json_data, list):
        # Itera sulle liste
        return [
            rimuovi_null(item)
            for item in json_data
            if item not in (None, 0, "", 0.0, [], {}) and rimuovi_null(item) != {}
        ]
    else:
        # Ritorna il valore se non è un dizionario o una lista
        return json_data

    
   

def leggi_excel(percorso_file_excel):
    # Leggi il file Excel
    CartellaTest=crea_cartella_con_contatore(direttivaBase,nomeCartellaTest)
    df = pd.read_excel(percorso_file_excel)

    conn=connessioneDataBase()
    # Itera attraverso le righe del DataFrame
    for index, row in df.iterrows():
        # Estrai i valori e gestisci i vuoti
        servizio = row['Servizio'] if pd.notna(row['Servizio']) else 'Valore di default'
        url = row['Url'] if pd.notna(row['Url']) else 'Valore di default'
        DBGruppo = row['DBGruppo'] if pd.notna(row['DBGruppo']) else 'Valore di default'
        corpo_json = row['CorpoJSON'] if pd.notna(row['CorpoJSON']) else {}
        ListaParametridaRimuovere = row['AttributiDaRimuovere'] if pd.notna(row['AttributiDaRimuovere']) else []
        NomeFile = row['NomeFile'] if pd.notna(row['NomeFile']) else 'Valore di default'
        if corpo_json:
            try:
                corpo_json = json.loads(row['CorpoJSON'])  # Converte la stringa JSON in un dizionario Python
            except json.JSONDecodeError as e:
                print(f"Errore nel parsing del JSON: {e}")
                corpo_json = {}
            
            if isinstance(ListaParametridaRimuovere, str) and len(ListaParametridaRimuovere) > 0:
                try:
                    ListaParametridaRimuovere = ast.literal_eval(ListaParametridaRimuovere)
                except (ValueError, SyntaxError):
                    print(f"Errore nel parsing della lista: {ListaParametridaRimuovere}")
                    ListaParametridaRimuovere = []
        

        
        sigleDoc = ["OPA", "OPV", "ORA", "ORP", "ORV", "RDA", "RDL"]
        global isCreaORVdaOPV
        global isCreaOPV
        if "/ES_ORV/App-Sistemi/Crea" in url and isCreaOPV:
            isCreaORVdaOPV=True
        else:
            isCreaORVdaOPV=False
        
        if "/ES_OPV/App-Sistemi/Crea" in url:
            isCreaOPV=True
        else:
            isCreaOPV=False
        
        #Questa parte funziona e permette al programma di capire quando si vuole creare un ordine a partire dal preventivo, non la implemento ancora in questa 
        #parte per non appesantire, bisogna poi capire quante concatenazioni sono possibili perchè per adesso ce n'è una sola e quindi si può mettere un if nel crea,
        # ma se domani dovesserpo essere di più sarebbeil caso di studiare un codice più pulito

        if isCreaORVdaOPV:
            print("Sto creando un ordine di vendita dopo aver creato un preventivo")
        # Controllo se la stringa contiene almeno uno dei valori
        global isDocumento
        if any(sigla in url for sigla in sigleDoc):
            isDocumento=True
        else:
            isDocumento=False
        
        if servizio == "CREA":
         order_id=crea(token,DBGruppo,url,corpo_json,PercorsoFile,NomeFile,conn)
         
        elif servizio == "LEGGI":
         leggi(token,DBGruppo,url,corpo_json,order_id,PercorsoFile,NomeFile)
        elif servizio == "AGGIORNA":
         aggiorna(token,DBGruppo,url,corpo_json,order_id,PercorsoFile,NomeFile,conn)
        
         
        elif servizio == "ELIMINA":
         elimina(token,DBGruppo,url,corpo_json,order_id,PercorsoFile,NomeFile)
        elif servizio == "NUOVA CARTELLA":
         PercorsoFile= creaNuovaCartella(CartellaTest, url)
    # Aggiungi altre condizioni per gli altri servizi
        else:
         print(f"Servizio '{servizio}' non riconosciuto.")
    confronta_file_con_benchmark(CartellaTest,cartellaBenchmark,direttivaWinMerge)

def confronta_file_con_benchmark(cartellaNuova, cartellaBenchmark, direttivaWinMerge):
    """
    Confronta due cartelle utilizzando WinMerge.
    
    :param cartella_nuova: Percorso della cartella con i file nuovi.
    :param cartella_benchmark: Percorso della cartella con i file di benchmark.
    :param winmerge_path: Percorso dell'eseguibile di WinMerge.
    """
    try:
        # Esegui il confronto delle cartelle con WinMerge
        subprocess.run([direttivaWinMerge, cartellaBenchmark, cartellaNuova], check=True)
        print(f"Confronto tra {cartellaNuova} e {cartellaBenchmark} completato.")
    except FileNotFoundError:
        print(f"Errore: WinMerge non trovato al percorso {direttivaWinMerge}")
    except subprocess.CalledProcessError as e:
        print(f"Errore durante il confronto: {e}")
    except Exception as e:
        print(f"Errore inaspettato: {e}")
     

    

# Utilizzo della funzione
base_url = "http://toinsxa0040:9050/api"
percorso_file_excel = 'c:\\Users\\fiore\\Documents\\TEST NON REGRESSIONE\\TEST\\ORV\\Cartella excel casistiche\\Casistiche.xlsx'  
dati_servizi = leggi_excel(percorso_file_excel)






