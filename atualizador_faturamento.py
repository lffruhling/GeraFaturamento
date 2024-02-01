#Interface
import PySimpleGUI as sg
#import psutil
import signal
import requests
import time
import wget
import os
from zipfile import ZipFile
import os

tela = None

def baixarArquivo():
    if os.path.isfile('C:/Temp/atualizacao_faturamento.zip'):
        os.remove('C:/Temp/atualizacao_faturamento.zip')
        
    wget.download('http://portugues.wiki.br/corretor/faturamento.zip', 'C:/Temp/faturamento.zip', bar=bar_custom)
    print('Download concluído')
    
    if os.path.isfile('C:/Temp/faturamento.zip'):
        print('Extraindo dados')
        sistema = ZipFile('C:/Temp/faturamento.zip', 'r')
        sistema.extractall('C:/Temp/')
        sistema.close()
        
        os.remove('C:/Temp/faturamento.zip')
    
def progress(valor):
    global tela
    progress_bar = tela['progressbar']    
    progress_bar.UpdateBar(valor)

def encerraCorretor():
    # PROCNAME = "main.exe"

    #for proc in psutil.process_iter():
    #    if proc.name() == PROCNAME:
    #        proc.kill()
    
    os.system('taskkill /im GeraFaturamento.exe')

def bar_custom(atual, total, largura=80): 
    global tela
    valor = int(atual/total*100)
    tela['titulo'].Update('Atualizando... Aguarde...')
    tela['status'].Update('Realizando download do arquivo... ' + str(valor) + '% Concluído')
    progress(valor) 

def telaAtualizador():       
    global tela
    sg.theme('Reddit')
    
    alcadas = [    
                [sg.Text(text='Nova Versão Disponível', text_color="Black", font=("Arial",12, "bold"), expand_x=True, justification='center', key='titulo')],
                [sg.Text(text='Deseja atualizar agora?', text_color="Black", font=("Arial",9, "bold"), expand_x=True, justification='center', key='status')],
                [sg.ProgressBar(100, orientation='h', size=(30, 10), key='progressbar', visible=True, expand_x=True)]                  

              ]
    
    tela = sg.Window('Atualização Disponível', alcadas)                

    while True:                
        eventos, valores = tela.read(timeout=0.1)

        encerraCorretor()
        baixarArquivo()
        return None
    tela['status'].Update('Aguarde... Abrindo Sistema de Faturamentos...')
    os.system(r"C:\Temp\GeraFaturamento\GeraFaturamento.exe")
    sg.popup_ok('Sistema de Faturamentos atualizado com sucesso!')
        
telaAtualizador()