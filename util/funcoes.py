import shutil
from asyncio.windows_events import NULL
#import pyautogui as p
import datetime
from datetime import datetime
from datetime import datetime
from datetime import date
import time
import os

from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import smtplib

import MySQLdb 
import json
#import urllib3
#import requests
import base64
#from loguru import logger as Logs
#import pyqrcode
from io import BytesIO
import csv

import pdfplumber

feriados = [date(2022,10,12), date(2022,11,2), date(2022,11,15), date(2022,12,25),
            date(2023,1,1), date(2023,12,25), date(2023,4,7), date(2023,4,21),
            date(2023,5,1), date(2023,9,7), date(2023,10,12), date(2023,11,2),
            date(2023,11,15), date(2023,12,25)]

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

def data_hoje(formato= "%d/%m/%Y %H:%M:%S"):
    now = datetime.now()
    return now.strftime(formato)

def formata_data_banco(data, formato_atual, novo_formato= "%Y-%m-%d %H:%M:%S"):
    if type(data) is str:
        data = datetime.strptime(data, formato_atual)

    return data.strftime(novo_formato)  
                
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