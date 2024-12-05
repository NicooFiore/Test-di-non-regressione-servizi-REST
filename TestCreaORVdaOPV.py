import requests
import json
import time
import os
import pandas as pd
import numpy as np
import ast

#base_url = "http://toinsxa0040:9050/api"
#Leggi 
token="ClLeZowW5n4hVK8Lgb/2mDfgoKeNx1pCeawdZh4t4EmXBnMx0yw5klKJVjMSQaGZlfZbr11dNWQxkaWb91GJPEXxMTg5TqWum6lnsZ+qaRDf2Yl/wumY3w=="
idDocumento=''
parametriDaRimuovereDoc=["Documento.IdDocumento","Documento.DataRegistrazione","Documento.ParametriStampa.ConfermaOrdine"]
#numRegistraz=''
ORVdaOPV=True
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

    
#token=creaToken()




def salvaSuFile(fileDascrivere,directory,risposta):
    if not os.path.exists(directory):
     os.makedirs(directory)

    # Crea il percorso completo del file
    percorso_completo = os.path.join(directory, fileDascrivere)

    # Apri il file in modalità scrittura e scrivi del testo
    with open(percorso_completo, 'w') as file:
       file.write(json.dumps(risposta, indent=4))


def crea(token,DBgruppo,url,payload,fileDascrivere,directory):
    url_crea = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }
    
    #print(f"Payload inviato a {url_crea}: {json.dumps(payload, indent=4)}")
    

# devo mettere un if che controlla se si vuole creare un orv a partire da un opv, se si è in questo caso dovrà prendere il riferimento all'ordine e modificarlo
    if "ORV" in url:
        payload=modificaRiferimRigaDaEvadere(payload)
    response = requests.post(url_crea, headers=headers, data=json.dumps(payload))
    #print(idDaSalvare)
    json_data=response.json()
    
    if response.status_code == 200:
        idDocumento = response.json().get("IdDocumento") 
        if "OPV" in url:
         global numRegistraz
         numRegistraz= response.json().get("Documento", {}).get("NumRegistraz")
         print(numRegistraz)
        print(f"Ordine creato con ID: {idDocumento}")
        json_data=rimuovi_attributi(json_data)
        json_data=rimuovi_null(json_data)
        salvaSuFile(fileDascrivere,directory,json_data)
        return idDocumento
    else:
        print(f"Errore nella creazione dell'ordine: {response.status_code}, {response.text}")
        salvaSuFile(fileDascrivere,directory,json_data)
        return None
    

def leggi(token,DBgruppo,url,payload,valoreId,fileDascrivere,directory):
    url_leggi = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }

    if "IdDocumento" in payload:
        print(payload)
        payload["IdDocumento"] =valoreId 
        print(payload)
        print(f"Modificata la chiave  'IdDocumento' con il valore '{valoreId}'.")
        response = requests.post(url_leggi, headers=headers, data=json.dumps(payload))
    else:
        print(f"La chiave 'IdDocumento' non esiste nel JSON.")
        response = requests.post(url_leggi, headers=headers, data=json.dumps(payload))
    
    
    #print(response.json())
    #time.sleep(5)

    if response.status_code == 200:
       json_data=response.json()
       rimuovi_attributi(json_data)
       json_data=rimuovi_null(json_data)
       salvaSuFile(fileDascrivere,directory,json_data)
       
        
    else:
        print(f"Errore nella creazione dell'ordine: {response.status_code}, {response.text}")
        salvaSuFile(fileDascrivere,directory,response.json())
        return None
    
def aggiorna(token,DBgruppo,url,payload,valoreId,fileDascrivere,directory):
    url_elimina = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }

    if "IdDocumento" in payload:
            payload["IdDocumento"] =valoreId 
            print(f"Modificata la chiave 'IdDocumento' con il valore '{valoreId}'.")
    else:
        print(f"La chiave 'IdDocumento' non esiste nel JSON.")
    
    response = requests.post(url_elimina, headers=headers, data=json.dumps(payload))

    if response.status_code == 200:
        print(f"Ordine aggiornato con successo")
        
    else:
        print(f"Errore nell'aggiornamento dell'ordine: {response.status_code}, {response.text}")
        
    
    salvaSuFile(fileDascrivere,directory,response.json())    

def elimina(token,DBgruppo,url,payload,valoreId,fileDascrivere,directory):
    url_elimina = base_url+url
    headers = {
        "X-BC-Authorization": f"{token}",
        "Content-Type": "application/json",
        "X-BC-Gruppo" : f"{DBgruppo}"
    }

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
        
    
    salvaSuFile(fileDascrivere,directory,response.json())


def rimuovi_attributi(json_data):
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
             #print(f"Removing {attributi[-1]} from {temp_data}")
             temp_data.pop(attributi[-1])
            elif isinstance(temp_data, list):  # Se l'elemento è una lista
                for item in temp_data:
                    if isinstance(item, dict) and attributi[-1] in item:
                        item.pop(attributi[-1])  # Rimuovi la chiave da ogni dizionario nella lista
            else:
                print(f"Chiave '{attributi[-1]}' non trovata in {temp_data}")

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

def modificaRiferimRigaDaEvadere(payload):
    
    if isinstance(payload, dict):
        # Controlla se esiste il campo 'Righe'
        if "Righe" in payload and isinstance(payload["Righe"], list):
            for riga in payload["Righe"]:
                if isinstance(riga, dict) and "RiferimRigaDaEvadere" in riga:
                    RiferimRigaDaEvadere=riga["RiferimRigaDaEvadere"]
                    parti = RiferimRigaDaEvadere.split(".")
                    parti[2] = str(numRegistraz)
                    parti=".".join(parti) 
                    riga["RiferimRigaDaEvadere"]=parti
                    
        else:
            # Ricorsivamente analizza i figli del dizionario
            for chiave, valore in payload.items():
                modificaRiferimRigaDaEvadere(valore)
    elif isinstance(payload, list):
        # Analizza ricorsivamente gli elementi della lista
        for elemento in payload:
            modificaRiferimRigaDaEvadere(elemento)
    return payload
    
    
   
   

def leggi_excel(percorso_file_excel):
    # Leggi il file Excel
    df = pd.read_excel(percorso_file_excel)


    # Itera attraverso le righe del DataFrame
    for index, row in df.iterrows():
        # Estrai i valori e gestisci i vuoti
        servizio = row['Servizio'] if pd.notna(row['Servizio']) else 'Valore di default'
        url = row['Url'] if pd.notna(row['Url']) else 'Valore di default'
        DBGruppo = row['DBGruppo'] if pd.notna(row['DBGruppo']) else 'Valore di default'
        corpo_json = row['CorpoJSON'] if pd.notna(row['CorpoJSON']) else {}
        PercorsoFile = row['PercorsoFile'] if pd.notna(row['PercorsoFile']) else 'Valore di default'
        NomeFile = row['NomeFile'] if pd.notna(row['NomeFile']) else 'Valore di default'
        try:
             corpo_json = json.loads(row['CorpoJSON'])  # Converte la stringa JSON in un dizionario Python
        except json.JSONDecodeError as e:
            print(f"Errore nel parsing del JSON: {e}")
            corpo_json = {}

        #print(idDsalvare)
        #print(corpo_json)

        if servizio == "CREA":
            # Chiamata alla funzione di creazione con i parametri necessari
         order_id=crea(token,DBGruppo,url,corpo_json,NomeFile,PercorsoFile)
         time.sleep(10)
         
        elif servizio == "LEGGI":
        # Chiamata alla funzione di lettura
         time.sleep(5)
         leggi(token,DBGruppo,url,corpo_json,order_id,NomeFile,PercorsoFile)
        elif servizio == "AGGIORNA":
        # Chiamata alla funzione di lettura
         time.sleep(5)
         aggiorna(token,DBGruppo,url,corpo_json,order_id,NomeFile,PercorsoFile) 
         
        elif servizio == "ELIMINA":
        # Chiamata alla funzione di lettura
         time.sleep(5)
         elimina(token,DBGruppo,url,corpo_json,order_id,NomeFile,PercorsoFile)
    # Aggiungi altre condizioni per gli altri servizi
        else:
         print(f"Servizio '{servizio}' non riconosciuto.")

def eseguiConfronti():
    print(True)       

    

# Utilizzo della funzione
base_url = "http://toinsxa0040:9050/api"
percorso_file_excel = 'c:\\Users\\fiore\\Documents\\TEST NON REGRESSIONE\\Servizi REST\\StrutturaFileCreaORVdaOPV.xlsx'  
dati_servizi = leggi_excel(percorso_file_excel)



