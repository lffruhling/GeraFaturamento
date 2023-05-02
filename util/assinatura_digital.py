import os
import json
#import dotenv  # required
#import requests
#from urllib3.util import Retry
#from loguru import logger as Logs
#from requests.adapters import HTTPAdapter

class AssinaturaDigital:
    def __init__(self):
        # pathEnv = 'F:\PROCESSOS E INOVAÇÃO\Python\envs\RPA_ASSINATURA_DIGITAL_CUSTEIO\.env'
        #Verficar se localizou meu env local para testes
        pathEnv = '../.env'
        dotenv.load_dotenv(dotenv.find_dotenv(pathEnv))
        self._header = json.loads(os.getenv('HEADER_PORTAL'))
        self._session = requests.Session()
        self._retry = Retry(connect=3, backoff_factor=2.5)
        self._adapter = HTTPAdapter(max_retries=self._retry)
        self._session.mount('http://', self._adapter)
        self._session.mount('https://', self._adapter)

    def enviaDocAssinatura(self, base64, name):
        try:
            body = {"fileName": name, "bytes": base64}
            response = self._session.post(f"{os.getenv('URL_API_PORTAL')}/upload", json=body, headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def criaAssinatura(self, sender, document, electronicSigners):
        try:
            body = {
                "document": document,
                "sender": sender,
                "signatureStandard": 1,
                "electronicSigners": electronicSigners
            }

            response = self._session.post(f"{os.getenv('URL_API_PORTAL')}/create", json=body, headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def criaAssinaturaLote(self, sender, tags, documents, electronicSigners, folderId):
        try:
            # body = {
            #         "sender": sender,
            #         "tags": tags,
            #         "signatureStandard": "CAdES",
            #         "electronicSigners": electronicSigners,
            #         "folderId": folderId,
            #         "batchName": pathName,
            #         "documents": documents
            #     }
            body = {
                    "sender": sender,
                    "tags": tags,
                    "signatureStandard": "CAdES",
                    "electronicSigners": electronicSigners,
                    "documents": documents,
                    "folderId": folderId,
                }

            response = self._session.post(f"{os.getenv('URL_API_PORTAL')}/createBatch", json=body, headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def idPastaPortal(self, nroAgencia):
        idAgencia = nroAgencia[-2:]
        if os.getenv('LOCAL').lower() == 'producao':
            if idAgencia == '02':
                return 198744
            elif idAgencia == '03':
                return 182033
            elif idAgencia == '04':
                return 198745
            elif idAgencia == '05':
                return 173257
            elif idAgencia == '06':
                return 194915
            elif idAgencia == '07':
                return 182035
            elif idAgencia == '08':
                return 198747
            elif idAgencia == '09':
                return 182036
            elif idAgencia == '10':
                return 194916
            elif idAgencia == '11':
                return 194917
            elif idAgencia == '12':
                return 198968
            elif idAgencia == '13':
                return 198754
            elif idAgencia == '14':
                return 182037
            elif idAgencia == '15':
                return 194918
            elif idAgencia == '16':
                return 199088
            elif idAgencia == '17':
                return 194919
            elif idAgencia == '18':
                return 199093
            elif idAgencia == '19':
                return 182038
            elif idAgencia == '20':
                return 198979
            elif idAgencia == '21':
                return 198973
            elif idAgencia == '22':
                return 199094
            elif idAgencia == '23':
                return 198758
            elif idAgencia == '24':
                return 182039
            elif idAgencia == '25':
                return 198977
            elif idAgencia == '26':
                return 182040
            elif idAgencia == '27':
                return 198760
            elif idAgencia == '28':
                return 198964
            elif idAgencia == '29':
                return 199095
            elif idAgencia == '30':
                return 182041
            elif idAgencia == '31':
                return 215561
            elif idAgencia == '32':
                return 194922
            elif idAgencia == '33':
                return 194923
            elif idAgencia == '34':
                return 250188
            elif idAgencia == '35':
                return 250577
            elif idAgencia == '36':
                return 322562
        else:
            if idAgencia == '01':
                return 2530
            elif idAgencia == '02':
                return 2517
            elif idAgencia == '03':
                return 2529