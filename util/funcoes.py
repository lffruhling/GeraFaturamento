import shutil
from asyncio.windows_events import NULL
#import pyautogui as p
import datetime
from datetime import datetime
from datetime import datetime
from datetime import date
import time
import os

#from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.common.by import By
#from selenium.webdriver.support.wait import WebDriverWait

from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import smtplib

#from cryptography.fernet import Fernet
import util.criptografia as c

import MySQLdb 
import json
#import urllib3
#import requests
import base64
#from loguru import logger as Logs
#import pyqrcode
from io import BytesIO
import csv

import util.constantes as cons
import pdfplumber



feriados = [date(2022,10,12), date(2022,11,2), date(2022,11,15), date(2022,12,25),
            date(2023,1,1), date(2023,12,25), date(2023,4,7), date(2023,4,21),
            date(2023,5,1), date(2023,9,7), date(2023,10,12), date(2023,11,2),
            date(2023,11,15), date(2023,12,25)]

 
v_path_criptografia = cons.PATH_CHAVE_CRIPTOGRAFIA

## Chave Criptografia das Credenciais ##
def recuperaCriptografia():           
    if os.path.isfile(v_path_criptografia):                
        retorno = open(v_path_criptografia, "rb").read()
    else:
        log('Gerando novas chaves de criptografia...\n')        
        chave = Fernet.generate_key()
        with open(v_path_criptografia, "wb") as arquivo_keys:
            arquivo_keys.write(chave)
        
        retorno = open(v_path_criptografia, "rb").read()
        
    return retorno

def criptografar(vtexto):
    mensagem     = vtexto.encode()
    criptografia = Fernet(recuperaCriptografia())
    
    retorno = criptografia.encrypt(mensagem) 
    
    return retorno.decode()

def descriptografar(vtexto):
    try:
        criptografia = Fernet(recuperaCriptografia())        
    except Exception as erro:
        log('Ocorreu um erro ao tentar recuperar a chave Mestre. ' + str(erro))
        
    try:        
        retorno = criptografia.decrypt(vtexto.encode())        
    except Exception as erro:
        log('Falha ao tentar descriptografar. Linha 71. Erro: ' + str(erro))
        
    return retorno.decode()
## Fim Criptografia Credenciais ##

## FUNÇÕES PARA NAVEGADOR WEB -- SELENIUM ##
def navegador_clica(navegador, vxpath, vdelay = 0, byname = False):
    try:
        if vdelay > 0:
                time.sleep(vdelay)
        vencontrou = False
        while not vencontrou:            
                if byname:
                    if navegador.find_element(By.NAME, vxpath) and navegador.find_element(By.NAME, vxpath).is_displayed():
                        vencontrou = True
                        navegador.find_element(By.NAME, vxpath).click()
                else:
                    if navegador.find_element(By.XPATH, vxpath) and navegador.find_element(By.XPATH, vxpath).is_displayed():
                        vencontrou = True
                        navegador.find_element(By.XPATH, vxpath).click()                                                        
    except:
        log('Não foi possível localizar o elemento para clicar: ' + vxpath)
        
def navegador_insere(navegador, vxpath, vtexto, byname = False):    
    try:
        vencontrou = False
        while not vencontrou:
            if byname:
                if navegador.find_element(By.NAME, vxpath):
                    vencontrou = True
                    elemento = navegador.find_element(By.NAME, vxpath)
                    elemento.click
                    elemento.clear
                    if elemento.is_selected:            
                        elemento.send_keys(vtexto)                    
                else:
                    print('Não localizou o elemento')
            else:
                if navegador.find_element(By.XPATH, vxpath):
                    vencontrou = True
                    elemento = navegador.find_element(By.XPATH, vxpath)
                    elemento.click
                    elemento.clear
                    if elemento.is_selected:            
                        elemento.send_keys(vtexto)                    
                else:
                    print('Não localizou o elemento')        
    except:
        log('Ocorreu uma falha ao tentar localizar o elemento: ' + vxpath)

def aguarda_download(diretorio, arquivo, tempo):
    for i in range(tempo):
        log('Aguardando Download: ' + str(tempo - i) + ' segundos restantes...')
        time.sleep(1)
        if any(arquivo in x and not '.crdownload' in x for x in os.listdir(diretorio)):
            return True
    else:
        return False
## FIM FUNÇÕES NAVEGADOR WEB -- SELENIUM ##

def moveArquivo(vNomeArquivo, vOrigem, vDestino):
    vRetorno = False
    try:
        if os.path.isfile(vOrigem + vNomeArquivo):
            try:
                pastaExiste(vDestino, True)
                shutil.copy2(vOrigem + vNomeArquivo, vDestino + '\\' + vNomeArquivo)
                log(vNomeArquivo + ' movido com sucesso!')
                vRetorno = True
            except Exception as erro:
                print(erro)
                log('Não foi possível mover o arquivo: ' + vOrigem + vNomeArquivo + ' para o destino: ' + vDestino + '\\' + vNomeArquivo)
        else:
            log('Não foi possível localizar o arquivo. ' + vOrigem + vNomeArquivo)
    except:      
        log('Falha ao tentar localizar o arquivo gerado. ')
    
    return vRetorno

def aguardaArquivo(vCaminhoCompleto):
    vRetorno = False

    while vRetorno == False:
        log('Aguardando disponibilizar o arquivo: ' + vCaminhoCompleto)
        time.sleep(1)
        if os.path.isfile(vCaminhoCompleto):
            vRetorno = True
    
    return vRetorno

def removeArquivo(pCaminhoArquivo):
    vRetorno = False

    try:
        os.remove(pCaminhoArquivo)
        vRetorno = True
    except Exception as e:
        log(f'Falha ao remover arquivo: {e}')

    return vRetorno


def log(msg):
    print(datetime.now().strftime("%m/%d/%Y %H:%M:%S") + ' => ' + msg + "\n")

def localizarImagem(vcaminho):         
    return p.locateOnScreen(vcaminho)

def encerraExe(exename) :
    os.system("taskkill /im "+exename+".exe")
    time.sleep(0.5)

def validaFimdeSemana(vData):
    if vData.weekday() >= 0 and vData.weekday() <= 4:
        return True
    else:
        return False

def validaFeriado(vData):
    vRetorno = True
    for d in feriados:        
        if vData.strftime("%Y/%m/%d") == d.strftime("%Y/%m/%d"):             
            vRetorno = False

    return vRetorno

def arquivoExiste(vCaminho):
    if os.path.isfile(vCaminho):
        return True
    else:
        return False

def envia_email(vusuario, vsenha, vassunto, vdestinatarios, vmensagem, vanexo=None, vQuebraLinha=True, vDestinatariosCopia=None):
    log('Envia email')    
    
    if vDestinatariosCopia != None:
        vdestinatarios = vdestinatarios + vDestinatariosCopia            
    
    msg             = MIMEMultipart()
    msg['From']     = vusuario
    msg['To']       = vdestinatarios
    #msg['To']      = ", ".join(vdestinatarios)
    if vDestinatariosCopia is not None:
        msg['Cc']   = ", ".join(vDestinatariosCopia)
    else:
        vDestinatariosCopia = []

    msg['Subject']  = vassunto

    if vQuebraLinha:
        vmensagem_formatada = str(vmensagem).replace("#", "</br>")
    else:
        vmensagem_formatada = vmensagem
    msg.attach(MIMEText(vmensagem_formatada, 'html'))
    
    ## Anexos
    if vanexo:
        part = MIMEBase('application', "zip")
        part.set_payload(open(vanexo, "rb").read())
        encoders.encode_base64(part)

        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(vanexo))
        msg.attach(part)

    smtp = smtplib.SMTP('smtp-exchange.sicredi.net', 25)
    smtp.starttls()

    smtp.login(vusuario, vsenha)
    smtp.sendmail(vusuario, vdestinatarios, msg.as_string())
    smtp.quit()
    
def ultimoDiaMesAnterior(vData:datetime):        
    vDia = vData.day        
    vMes = vData.month
    vAno = vData.year    
    
    if vMes > 1:
        vMes = vMes - 1
    else:
        vMes = 12
        vAno = vAno - 1
        
    if vMes in (1,3,5,7,8,10,12):
        vDia = 31
    elif vMes in (4,6,9,11):
        vDia = 30
    elif vMes == 2:
        ##Testa se é Bissesto
        if vAno in (2024,2028,2032): 
            vDia = 29
        else:
            vDia = 28
    
    #print(str(vDia)+'/'+str(vMes)+'/'+str(vAno))                
    return datetime(vAno, vMes, vDia)    

def primeiroDiaMesAnterior(vData:datetime):        
    vDia = vData.day        
    vMes = vData.month
    vAno = vData.year    
    
    if vMes > 1:
        vMes = vMes - 1
    else:
        vMes = 12
        vAno = vAno - 1
        
    vDia = 1    
        
    #print(str(vDia)+'/'+str(vMes)+'/'+str(vAno))                
    return datetime(vAno, vMes, vDia)   
    
## Credenciais inserção e Recuperação de Dados

def credenciaisDB():  
    try:                                        
        return MySQLdb.connect(host='10.4.21.24', database='credenciais', user=cons.DB_CREDENCIAIS_USUARIO, password=cons.DB_CREDENCIAIS_SENHA)           
    except Exception as erro:
        print('Ocorreu um erro ao tentar conectar ao banco de dados. ' + str(erro))
        with open("C:\\Temp\\SisPass_ErroConexaoDB_log.txt", "w") as arquivo:
            arquivo.write(str(erro))

def DB_0313():  
    try:                                        
        return MySQLdb.connect(host='10.4.21.24', database='0313', user=cons.DB_CREDENCIAIS_USUARIO, password=cons.DB_CREDENCIAIS_SENHA)           
    except Exception as erro:
        print('Ocorreu um erro ao tentar conectar ao banco de dados. ' + str(erro))
        with open("C:\\Temp\\DB0313_ErroConexao_log.txt", "w") as arquivo:
            arquivo.write(str(erro))

def recuperaCredencial(vnome):
  try:
    print('Buscando credencial ' + str(vnome) + '\n')  
    
    try:
        conexao = credenciaisDB()
        cursor  = conexao.cursor()  
    except Exception as erro:
        print('Falha ao tentar conectar ao Banco de Dados das credenciais. ' + str(erro))
    
    vparam = [(vnome.upper())]    
    
    cursor.execute("SELECT usuario, senha, descricao FROM credenciais where upper(credenciais.nome) = %s", vparam)
    resultado = cursor.fetchall()    
        
    if resultado != ():        
        usuario   = str(resultado[0][0])
        senha     = str(resultado[0][1])
        descricao = str(resultado[0][2])
    
        print('Usuário encontrado: ' + usuario)
        print('Descrição: ' + descricao + ' \n')
        
        try:
            senha = descriptografar(senha)            
            resultado = {
                            'usuario'   : usuario,
                            'senha'     : senha,
                            'descricao' : descricao
                         }
        except:            
            print('Ocorreu um erro ao tentar descriptografar a senha da credencial recuperada.')
                    
    conexao.close()    
    
    return resultado
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar recuperar as credenciais no DB. ' + str(erro))
    return False
    
def consultaDB(expressao):
  try:
    
    conexao = credenciaisDB()
    cursor = conexao.cursor()
    
    cursor.execute(expressao)
    resultado = cursor.fetchall()
    
    cursor.close()
    conexao.close()
    
    return resultado
        
  except Exception as erro:
    print('Falha ao tentar executar consulta ao DB. ' + str(erro))
    
def updateSQL(vsql):
  try:  
    conexao = credenciaisDB()
    cursor = conexao.cursor()  
    
    cursor.execute(vsql)    
    conexao.commit()
    
    cursor.close()
    conexao.close()
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar executar a instrução SQL. ' + str(erro))
    
def insertSQL(vsql,vparametros):
  try:  
    conexao = credenciaisDB()
    cursor = conexao.cursor()  
    
    cursor.execute(vsql,vparametros)    
    conexao.commit()
    
    cursor.close()
    conexao.close()
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar executar a instrução SQL. ' + str(erro))
    
def consultaDBapi(expressao):
  try:
    
    conexao = credenciaisDB()
    cursor = conexao.cursor()
    
    cursor.execute(expressao)
    row_headers=[x[0] for x in cursor.description]
    resultado = cursor.fetchall()
    
    vjson = []
    for linha in resultado:
        vjson.append(dict(zip(row_headers,linha)))
    
    cursor.close()
    conexao.close()
    
    return vjson
    return json.dumps(vjson)
        
  except Exception as erro:
    print('Falha ao tentar executar consulta ao DB. ' + str(erro))
    
def arquivoExiste(vLocalArquivo):
    return os.path.isfile(vLocalArquivo)

def pastaExiste(vLocal, vCriarDiretorio=False):
    if vCriarDiretorio:
        try:
            if os.path.isdir(vLocal):
                return True
            else:
                os.makedirs(vLocal)
                return True
        except Exception as e:
            log(e)
            return False
    else:
        return os.path.isdir(vLocal)

def gravalog(nome_automacao, msg, is_erro=True, db=False):
    try:
        if is_erro:
            vtipo = 'ERRO'
            vmensagem = f'\n{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}    ERRO: {str(msg)}'
        else:
            vtipo = 'INFO'
            vmensagem = f'\n{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}    INFO: {str(msg)}'
            
        vNome    = datetime.now().strftime("%d-%m-%Y") + '_log.txt'
        vLocal   = f'C:\Temp\{str(nome_automacao).upper()}'
        pastaExiste(vLocal, True)
        vArquivo = f'{vLocal}\{vNome}'

        if os.path.isfile(vArquivo):
            with open(vArquivo, 'a') as log:
                log.write(vmensagem)
        else:
            with open(vArquivo, 'w') as log:
                log.write(vmensagem)
                
        if db:
            try:
                conexao = DB_0313()
                cursor  = conexao.cursor()  
                                                                
                vparametros = (str(datetime.now().strftime("%Y-%m-%d %H:%M:%S")), str(vtipo), str(nome_automacao), str(msg))
                
                cursor.execute('INSERT INTO logs(data_hora,tipo,sistema,descricao) VALUES(%s,%s,%s,%s)',vparametros)    
                conexao.commit()
    
                cursor.close()
                conexao.close()  
            except:
                print('Falha ao tentar gravar log no Db')
    
        print(vmensagem)
    except Exception as erro:
        log('Falha ao tentar gravar log na pasta Temp. ' + str(erro)) 
        
def verificaFeriado():
  try:
    vdata = datetime.now()
    print('Verificando se a data ' + vdata.strftime("%d/%m/%Y") + ' é um feriado. \n')  
    
    try:
        conexao = DB_0313()
        cursor  = conexao.cursor()  
    except Exception as erro:
        print('Falha ao tentar conectar ao Banco de Dados 0313. ' + str(erro))
    
    vparam = [(vdata.strftime("%Y-%m-%d"))]    
    
    cursor.execute("SELECT descricao FROM feriados where data = %s", vparam)
    resultado = cursor.fetchall()    
        
    if resultado != ():
        vResultado = True
    else:
        vResultado = False   
                    
    conexao.close()    
    
    return vResultado
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar recuperar os feriados no DB 0313. ' + str(erro))
    return False      

def consulta_api_denodo_rest(nome_db, nome_view, query='', fields='', timeout=20):
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    credenciais = recuperaCredencial('Denodo')
    vusuario = str(credenciais['usuario'])
    vsenha   = str(credenciais['senha'])
        
    headers = urllib3.make_headers(basic_auth='' + vusuario + ':' + vsenha + '')
    headers.update({'Content-Type': 'application/json'})
    
    url = 'https://virtualizador.sicredi.net/denodo-restfulws/' + nome_db + '/' \
            'views/' + nome_view + '?' \
            + ('$select=' + fields + '' if fields else '') \
            + ('&' + query + '' if query else '') +\
            '&$format=JSON'        
            
    #print('\n'+url+'\n\n')
    r = requests.get(url, verify=False, headers=headers, timeout=timeout)
    if r.status_code == 200:
        return r.json()
    else:
        raise Exception("Erro de consulta à API (" + str(r.status_code) + "): " + r.url)

def data_hoje(formato= "%d/%m/%Y %H:%M:%S"):
    now = datetime.now()
    return now.strftime(formato)

def formata_data_banco(data, formato_atual, novo_formato= "%Y-%m-%d %H:%M:%S"):
    if type(data) is str:
        data = datetime.strptime(data, formato_atual)

    return data.strftime(novo_formato)

def consultaDB0313(nome_tabela, fields, filtros = [], order_by='', group_by=''):
    resultado = None
    vContinua = True
    try:
        conexao = DB_0313()
        cursor  = conexao.cursor()
    except Exception as erro:
        print('Ocorreu um erro ao tentat conectar ao Banco de Dados 0313. ' + str(erro))
        vContinua = False
        time.sleep(60)

    if vContinua:
        vsql = 'SELECT ' + fields + ' FROM ' + nome_tabela + ' WHERE 1=1'

        if filtros != []:
            for filtro in filtros:
                vsql = vsql + ' AND ' + str(filtro)
                
        if group_by != '':
            vsql = vsql + ' GROUP BY ' + group_by

        if order_by != '':
            vsql = vsql + ' ORDER BY ' + order_by

        #print('\nInstrução SQL: ' + vsql + '\n')
        try:
            cursor.execute(vsql)
            resultado = cursor.fetchall()
        except Exception as erro:            
            print('Ocorreu um erro ao tentar executar a consulta. ' + str(erro))
            print('\nInstrução SQL: ' + vsql + '\n')
            vContinua = False

    return resultado

def insereDB0313(nome_tabela, fields, values, mostaLog = False):
    resultado = None
    vContinua = True
    try:
        conexao = DB_0313()
        cursor  = conexao.cursor()
    except Exception as erro:
        print('Ocorreu um erro ao tentar conectar ao Banco de Dados 0313. ' + str(erro))
        vContinua = False
        time.sleep(60)

    if vContinua:
        vsql = 'INSERT INTO ' + nome_tabela + '(' + fields + ') VALUES('+ values + ')'

        if mostaLog:
            print('\nInstrução INSERT: ' + vsql + '\n')

        try:
            cursor.execute(vsql)
            conexao.commit()

            resultado = cursor.rowcount
        except Exception as erro:
            print('Ocorreu um erro ao tentar executar a consulta. ' + str(erro))

    return resultado

def updateDB0313(nome_tabela, _set, _where):
    resultado = None
    vContinua = True
    try:
        conexao = DB_0313()
        cursor  = conexao.cursor()
    except Exception as erro:
        print('Ocorreu um erro ao tentat conectar ao Banco de Dados 0313. ' + str(erro))
        vContinua = False
        time.sleep(60)

    if vContinua:
        vsql = 'UPDATE ' + nome_tabela + ' SET ' + _set + ' WHERE ' + _where

        print('\nInstrução UPDATE: ' + vsql + '\n')
        try:
            cursor.execute(vsql)
            conexao.commit()

            resultado = cursor.rowcount
        except Exception as erro:
            print('Ocorreu um erro ao tentar executar o update. ' + str(erro))

    return resultado

def carregaParametro(nome_automacao, id_parametro):
    vContinua = True
    valor     = None
    tipo      = None
    
    try:
        conexao = DB_0313()
        cursor  = conexao.cursor()                
    except Exception as erro:
        gravalog(nome_automacao, 'Ocorreu um erro ao tentar conectar ao Banco de Dados 0313. ' + str(erro), True, False)
        vContinua = False
        
    if vContinua:
        vSQL = 'SELECT valor, tipo from parametros_automacoes WHERE id = ' + str(id_parametro)
        
        cursor.execute(vSQL)
        resultado = cursor.fetchall()
        
        if len(resultado) > 0:
            tipo            = resultado[0][1]
            valor_parametro = resultado[0][0]
            
            if tipo == 'Int':
                valor = int(valor_parametro)
            if tipo == 'String':
                valor = str(valor_parametro)
            if tipo == 'Boolean':
                if str(valor_parametro) == 'True':
                    valor = True
                else:
                    valor = False

    return valor                        
                
def getExtensaoUrl(url):
    urlSplit = url.split("/")
    urlSplit = urlSplit[-1].split("?")
    urlSplit = urlSplit[0]
    urlSplit = urlSplit.split(".")
    return urlSplit[-1]

def formataNomeAnexo(url, nome):
    return nome + "." + getExtensaoUrl(url)

def verificaDiretorios(path):
    if not os.path.exists(path):
        os.makedirs(path)

def deletaDiretorio(path):
    try:
        shutil.rmtree(path)
    except Exception as e:
        Logs.error(e)

def converteBase64(pathFile):
    try:
        with open(pathFile, "rb") as pdf_file:
            return base64.b64encode(pdf_file.read()).decode('utf-8')
    except Exception as e:
        Logs.error(e)

def retornaCPFColaborador(email):
    pathColabs = 'F:\PROCESSOS E INOVAÇÃO\Python\envs\RPA_ASSINATURA_DIGITAL_CUSTEIO\colaboradores.csv'
    # pathColabs = 'colaboradores.csv'
    try:
        with open(pathColabs, 'r', encoding="utf8") as file:
            file_csv = csv.reader(file, delimiter=';')
            for row in file_csv:
                if email in row:
                    return str(row).split(",")[5].replace("'", "").strip().zfill(11)
            else:
                return ''
    except Exception as e:
        Logs.error(e)

def geraQrCode(url):
    try:
        c = pyqrcode.create(url)
        s = BytesIO()
        c.png(s, scale=6)
        encoded = base64.b64encode(s.getvalue()).decode("utf-8")
        #return base64.b64encode(img).decode('utf-8')
        return encoded

    except Exception as e:
        Logs.error(e)


def converterPDF(vArquivo):
    try:
        ## Carrega arquivo
        pdf = pdfplumber.open(vArquivo)
        vArquivoTxt = vArquivo.replace("pdf", 'txt')
        ## Converte PDF em TXT
        for pagina in pdf.pages:
            texto = pagina.extract_text(x_tolerance=1)
            with open(vArquivoTxt, 'a') as arquivo_txt:
                arquivo_txt.write(str(texto))
            arquivo_txt.close()


        pdf.close()

        return True

    except Exception as erro:
        with open('c:\\Temp\\Logs\\ERRO_LOG.txt', 'a') as arquivo_txt:
            arquivo_txt.write(str(erro))

        return False