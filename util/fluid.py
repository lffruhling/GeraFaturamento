#import dotenv  # required
import json
import os
#import requests  # required
#import wget  # required
#import util.constantes as const
#from requests.adapters import HTTPAdapter
#from urllib3.util import Retry  # required
#from loguru import logger as Logs

class Fluid:

    def __init__(self, _idUsuario, _idEmpresa, localEnv):
        # Localiza o Caminho do .env
        # pathEnv = f'F:\PROCESSOS E INOVAÇÃO\Python\envs\{localEnv}\.env'
        pathEnv = f'..\.env'
        dotenv.load_dotenv(dotenv.find_dotenv(pathEnv))
        self._header = json.loads(os.getenv('HEADER_FLUID'))
        self._headerAnexo = json.loads(os.getenv('HEADER_FLUID_ANEXO'))

        self._idUsuario = _idUsuario
        self._idEmpresa = _idEmpresa
        self._session = requests.Session()
        self._retry = Retry(connect=3, backoff_factor=2.5)
        self._adapter = HTTPAdapter(max_retries=self._retry)
        self._session.mount('http://', self._adapter)
        self._session.mount('https://', self._adapter)

    def getFilaProcessos(self):
        try:
            response = self._session.get(
                f"{os.getenv('URL_API_FLUID')}processos/inbox/{self._idUsuario}/{self._idEmpresa}",
                headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def getProcesso(self, idProcesso):
        try:
            response = self._session.get(
                f"{os.getenv('URL_API_FLUID')}processos/visualizar/{idProcesso}/{self._idUsuario}",
                headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def getUrlAnexoFluid(self, idProcesso, hashs):
        try:
            body = {"processo": idProcesso, "hashs": hashs}
            response = self._session.post(f"{os.getenv('URL_API_FLUID')}processos/download", json=body,
                                          headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def baixaArquivo(self, url, nome):
        try:
            response = wget.download(url, nome)
            return response
        except Exception as e:
            Logs.error(e)

    def preencheCpoForm(self, idProcesso, idCampo, valorCpo, idNodoDestino):
        try:
            body = {
                "processo": idProcesso,
                "usuario": self._idUsuario,
                "formulario": idCampo,
                "valor": valorCpo,
                "destino": idNodoDestino
            }

            response = self._session.post(f"{os.getenv('URL_API_FLUID')}processos/salvar", json=body,
                                          headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def anexarArquivo(self, idProcesso, arquivo, idTipoDocumento, nroTitulo = None):
        try:
            nomeArquivo = os.path.basename(arquivo)
            pathArquivo = os.path.dirname(arquivo)
            if len(nomeArquivo) < 15:
                if nroTitulo != None:
                    novoNomeArquivo = nomeArquivo.split('.')[0] + "_" + str(nroTitulo).replace("-", "_")
                else:
                    novoNomeArquivo = nomeArquivo.split('.')[0]
                    
                novoNomeArquivo = novoNomeArquivo + "." + nomeArquivo.split('.')[1]
                os.rename(arquivo, pathArquivo+"/"+novoNomeArquivo)
                arquivo = pathArquivo + "/" + novoNomeArquivo
                nomeArquivo = novoNomeArquivo


            body = {
                "processo": idProcesso,
                "arquivo": nomeArquivo,
                "tipo": idTipoDocumento
            }

            rFluid = self._session.post(f"{os.getenv('URL_API_FLUID_ANEXO')}processo/anexar", json=body,
                                        headers=self._headerAnexo)
            if rFluid.status_code != 200:
                Logs.info(body)
                Logs.error(rFluid.status_code)
                Logs.error(rFluid.json())
                return

            documento = open(f'{arquivo}', 'rb')
            params = rFluid.json()['dados']['fields']
            params["file"] = documento
            url = rFluid.json()['dados']['url']

            rS3 = self._session.post(url, files=params)

            if rS3.status_code != 204:
                Logs.info(params)
                Logs.error(rS3.status_code)
                Logs.error(rS3.content)
                return

            documento.close()

            resultAnexo = {
                "fluid_status": rFluid.status_code,
                "fluid_json": rFluid.json(),
                "s3_status": rS3.status_code,
                "s3_json": rS3.content
            }

            Logs.info(resultAnexo)

            return resultAnexo

        except Exception as e:
            Logs.error(e)

    def protocolarProcesso(self, idProcesso, textoParecer, idNodoDestino, acao=const.PROTOCOLAR):
        try:
            body = {
                "acao": acao,
                "processo": idProcesso,
                "destino": idNodoDestino,
                "parecer": textoParecer,
                "parecer_restrito": 0,
                "usuario": self._idUsuario,
                "empresa_destino": self._idEmpresa,
                "usuario_destino": self._idUsuario
            }

            response = self._session.post(f"{os.getenv('URL_API_FLUID')}processos/protocolar", json=body,
                                          headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def abrirProcesso(self, tipoProcesso, tempo, responsavelCriacao, responsavelOrigem, responsavelDestino, empresaOrigem, empresaDestino, versaoProcesso):
        try:
            body = {
                "tipo_processo": tipoProcesso,
                "tempo": tempo,
                "responsavel_criacao": responsavelCriacao,
                "responsavel_origem": responsavelOrigem,
                "responsavel_destino": responsavelDestino,
                "empresa_origem": empresaOrigem,
                "empresa_destino": empresaDestino,
                "versao_processo": versaoProcesso
            }

            response = self._session.post(f"{os.getenv('URL_API_FLUID')}processos/novo", json=body,
                                          headers=self._header)
            return response
        except Exception as e:
            Logs.error(e)

    def getValorCampo(self, atributos: list, idCampo: int):
        try:
            for campo in atributos:
                if campo['id'] == idCampo:
                    return campo['valor']
        except Exception as e:
            Logs.error(e)

    def retornaDestinoAcao(self, acoes, destino):
        for acao in acoes:
            if destino in acao['descricao']:
                return acao['destino']