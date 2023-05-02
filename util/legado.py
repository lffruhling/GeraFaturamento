from datetime import datetime, timedelta
from email.mime import application
from genericpath import isdir
import ctypes
import time
#import pyautogui as p
import re
#import clipboard
import shutil
import os
import util.funcoes as f
from asyncio.windows_events import NULL

vmapaAcc = ['CCRC', 'CCRF', 'DCNA', 'DOCTOR', 'SAC', 'SACG', 'SACGR', 'SIAC', 'SIACRL', 'SIAT', 'SIATLB', 'SIATRL'] 
usuario  = os.getlogin()

def iniciaAcc(vnovo):
    try:    
        vAvanca = False

        ## Define caminho dos atalho do ACClient
        if vnovo:
            linkapp   = "C:" + os.environ["HOMEPATH"] + "\\Desktop\\AC Client.lnk"         
            link2app  = "C:\\Users\\Public\\Desktop\\AC Client.lnk"            
        else:       
            linkapp   = "C:" + os.environ["HOMEPATH"] + "\\Desktop\\ACClient (App-V).lnk"
            link2app  = "C:\\Users\\Public\\Desktop\\ACClient (App-V).lnk"            
            
        ## Tenta startar o arquivo link
        try:        
            time.sleep(2)               
            os.startfile(linkapp, 'open')            
            vAvanca = True            
        except:              
            try:    
                ## Se deu Except no primeiro, tenta o segundo Link
                time.sleep(1)                      
                os.startfile(link2app, 'open')
                vAvanca = True            
            except:      
                print('Falha ao tentar abrir o sistema Legado.')
                vAvanca = False    
                
        return True                           
    except Exception as erro:
        print('Ocorreu um erro ao tentar executar a aplicação. ' + str(erro))

def abre_legado(vsistema, vusuario, vsenha, vnovo = False):       
    f.log('Abrir Sistema ACCliente') 

    vAvanca = False

    while vAvanca == False:

        ## Inicia o ACClient e valida se Conectou ao Servidor
        if iniciaAcc(vnovo):
            # Valida se estabeleceu conexão, senão encerra o exe e abre novamente.            
            time.sleep(5)                        
            try:                    
                vAcc = f.localizarImagem("F:\\PROCESSOS E INOVAÇÃO\\Python\\Imagens\\ConexaoEstabelecida.PNG")
                #vAcc = f.localizarImagem(".\\Imagens\\ConexaoEstabelecida.PNG")                                                                        
                ## Por conta de um erro da lib pyscreeze, relacionado com caracteres acentuados no caminho das pastas, tive que copiar a Imagem para o Temp para poder utilizar ela
                #shutil.copyfile("F:\\PROCESSOS E INOVAÇÃO\\Python\\Imagens\\ConexaoEstabelecida.PNG", "C:\\Temp\\ConexaoEstabelecida.PNG")
                #path = "C:\\Temp\\ConexaoEstabelecida.PNG"
                #vAcc = f.localizarImagem(path)                                                                    
            except Exception as erro:
                f.log('Erro ' + str(erro))
                with open('C:\\Temp\\ERRO_RPA_Abre_Legado.txt', 'w') as arq:
                    arq.write('Não localizou a imagem de Conexão estabelecida.\n\nVerifique nas libs instaladas o pyscreeze, provavel problema em decodificar caracteres acentuados.\n\nhttps://stackoverflow.com/questions/71229356/failed-to-read-images-when-folder-has-special-characters-on-name \n\n ' + str(erro))
            if vAcc != None:
                p.click(vAcc)
                vAvanca = True
            else:
                f.encerraExe('ACClient')
                time.sleep(2)            

    if vAvanca:
        f.log('Opção de Sistema para Selecionar: ' + vsistema)             
        time.sleep(2)

        ## Da um tab na cooperativa e vai pra lista de sistemas
        ## Na lista de sistemas vai para o topo, para respeitar a ordem que foi definida na variável vmapaAcc        
        p.press("tab")
        time.sleep(2)
        p.press("home")

        ## Faz um laço nas opções até chegar no sistema da variavel "mapa" até chegar na opção que deve selecionar.        
        for opcao in vmapaAcc:
            f.log('Buscando opção: ' + vsistema + ' | Opção atual: ' + str(opcao))
            if opcao == vsistema:
                p.hotkey("alt", "e")
                time.sleep(5)

                user32 = ctypes.WinDLL('user32')
                hwnd = user32.GetForegroundWindow() 
                user32.ShowWindow(hwnd, 3)
                
                break
            else:
                p.press("down")
                time.sleep(0.2)

        try:
            vAvanca = False
            ## Localiza qual a posição referência da tela para poder manipular o Legado de acordo com coordenadas.            
            f.log('Capturar posição da tela.')
            
            try:                
                pegarPosicaoLegado("F:\\PROCESSOS E INOVAÇÃO\\Python\\Imagens\\login.png")
                f.log('Posição do Legado identificada com sucesso')
                #pegarPosicaoLegado(".\\Imagens\\login.png")
                #pegarPosicaoLegado("F:\\PROCESSOS E INOVAÇÃO\\Python\\Imagens\\login.png")                
                #shutil.copyfile("F:\\PROCESSOS E INOVAÇÃO\\Python\\Imagens\\login.PNG", "C:\\Temp\\login.PNG")
                #path = "C:\\Temp\\login.PNG"
                #pegarPosicaoLegado(path)                
            except Exception as erro:
                f.log('Falha ao tentar localizar posição do Legado. ' + str(erro))
                with open('C:\\Temp\\ERRO_RPA_Abre_Legado.txt', 'w') as arq:
                    arq.write('Não localizou a posição do Legado.\n\nVerifique nas libs instaladas o pyscreeze, provavel problema em decodificar caracteres acentuados.\n\nhttps://stackoverflow.com/questions/71229356/failed-to-read-images-when-folder-has-special-characters-on-name \n\n ' + str(erro))
            
            ## Aguarda carregar a tela de login
            try:
                if aguardaTexto('Login',3,1,False) == 0:
                    vAvanca = True

                    f.log('Insere as credenciais de acesso ao ' + vsistema)
                    p.write(vusuario)
                    p.press('enter')                    
                    time.sleep(1)
                    p.write(vsenha)
                    p.press('enter')
            except Exception as erro:
                f.log('Erro - Não localizou a tela de Login do ' + vsistema + '. ' + str(erro))            

            if vAvanca == True:
                
                ## Procura e testa se houve falha no Login
                time.sleep(2)
                vLogin = procurarTexto('Erro-0006')
                if vLogin != None:                
                    f.log('Falha de login no Sistema: ' + vsistema + '. Usuário utilizado: ' + vusuario)
                    f.log('O ACClient será encerrado. ')
                    time.sleep(2)
                    f.encerraExe('ACClient')
                else:
                    f.log('Login Ok')
                    
                    ## Procura se apareceu tela de selecionar a agência
                    vRetorno = procurarTexto('Usuários vinculados a')
                    if vRetorno != None:
                        escreveTexto('01')
                        time.sleep(2)
                        p.press('enter')
                
                    ## Procura se chegou ao Menu Inicial do sistema
                    time.sleep(1)
                    if aguardaTexto('Mensagens', 5) == 0:            
                        f.log('Menu do Sistema ' + vsistema)
                        return True                                      
        except Exception as verro:
            print('Falha ao tentar pegar posição da tela. ' + str(verro))   
            with open('C:\\Temp\\ERRO_RPA_Abre_Legado.txt', 'w') as arq:
                arq.write('Falha ao tentar pegar posição da tela. ' + str(erro))         

def escreveTexto(vtexto, intervalo = 0.4):
    p.typewrite(vtexto, interval=intervalo)

def pegarPosicaoLegado(vimagem):    
    i = 0
    vavanca = False

    while ((vavanca is False)and(i < 50)):
        ## Variavel global da posição da tela
        global vposicao
        
        vposicao   = None
        p.FAILSAFE = False
        p.moveTo(0, 0)    
        #p.pyscreeze.useOpenCV = True        
        #pyscreeze.useOpenCV = True
        vposicao = p.locateOnScreen(vimagem, grayscale=True, confidence=0.9)
        #vposicao = pyscreeze.locateOnScreen(vimagem, grayscale=True, confidence=0.9)        

        if (vposicao != None):            
            vavanca = True
            f.log('Tela Legado localizada em => ' + str(vposicao))
        else:            
            f.log('Posição não localizada. Tentativa: ' + str(i))
            time.sleep(2)
            i = i + 1

    return vposicao
            
def pegarTextoLegado():    
    try:        
        clipboard.copy('')
        #log('Capturando Tela legado: ' + str(vposicao[0]) + ', ' + str(vposicao[1]))
        ## Clica no canto superior esquerdo da tela Legado
        p.click(vposicao[0] + 4, vposicao[1] + 4)
        
        ## Clica no canto inferior direito da tela Legado
        ## Selecionando e copiando todo conteúdo
        p.keyDown('shift')
        p.click((vposicao[0]+4) + (vposicao[2] - 10), (vposicao[1] + 4) + (vposicao[3] - 6))        

        p.keyUp('shift')        
        p.press('enter')
    except:
        print('Falha ao tentar copiar tela Legado. ')
    
    return clipboard.paste()

def procurarTexto(vtexto):
    f.log('Procurando expressão no legado: ' + vtexto)
    vcompilado = re.compile(vtexto)
    result     = vcompilado.search(pegarTextoLegado())
    
    if result is not None:
        f.log('Resultado encontrado: ' + str(result))        
        return [result.start(), result.end()]    
    else:
        return None
    
def procurarTextoLinha(vtexto, linha=0, tentativas=1):    
    vcompilado = re.compile(vtexto)    
    vEncontrou = False
    
    if linha > 0:
        for x in range(tentativas):                 
            i = 0                    
            try:               
                with open('C:\\Temp\\legado.txt', 'w', encoding="utf-8") as tela:
                    vlista = pegarTextoLegado().splitlines()
                    for v_linha in vlista:                    
                        tela.write(v_linha+'\n')
            except Exception as erro:
                print('Ocorreu um erro ao tentar copiar o texto para o arquivo. ' + str(erro))
                    
            try:                                        
                with open('C:\\Temp\\legado.txt', 'r', encoding="utf-8") as legado:                
                    vlinhas = legado.readlines()
                                                    
                    for l in vlinhas:                    
                        i = i + 1
                        if vtexto in l:                            
                            if linha == i:
                                vEncontrou = True                                
                                break                                
            except Exception as erro:
                print('Ocorreu um erro ao tentar percorrer as linhas. ' + str(erro))
                  
            if vEncontrou:
                break
            
            if (tentativas > 1):
                time.sleep(1.5)                            
                
        return vEncontrou
                        
    else:
        result     = vcompilado.search(pegarTextoLegado())
        if result is not None:
            f.log('Resultado encontrado: ' + str(result))        
            return [result.start(), result.end()]    
        else:
            return None
    
def retornaTextoLinha(linha = 0):
    vlista = pegarTextoLegado().splitlines()
    vCount = 0
    for v_linha in vlista:
        if vCount == linha:
            return v_linha
        vCount = vCount + 1


def aguardaTexto(vtexto, vtentativas = 20, intervalo = 1.5, vprint = False):
    f.log('Aguarda expressão no legado: ' + vtexto)
    i = 0
    while True:
        time.sleep(intervalo)
        retorno = NULL
        retorno = procurarTexto(vtexto)
        
        if vprint:
            print(retorno)
       
        ## Se localizou
        #if retorno[0] != None and retorno[1] != None:
        if retorno != None:
            return 0
        else:
            i = i + 1
            if i >= vtentativas:
                return None
            
def enviaTecla(vTecla):
    p.press(vTecla)
    
def encerraSistema():
    while procurarTexto('deseja abandonar o Sistema') == None:
        enviaTecla('esc')
        time.sleep(2)
    
    enviaTecla('S')
    time.sleep(2)

    return True

def printPosition(vPosicao, vDestino, vNome):    
    f.log('Screenshot')
    try:
        vImagem = p.screenshot(vDestino + '\\' + vNome + '.png', region=(vPosicao))
    except Exception as erro:
        f.log('Ocorreu um erro ao tentar capturar o Screenshot. ' + str(erro))

def retornaMenuPrincipal():
    pressEsc()
    while aguardaTexto('Retorna ao Menu Anterior') == 0:
        pressEsc()
        if aguardaTexto('Retorna ao Sistema Operacional', vtentativas=2) == 0:
            break

def pressEnter():
    enviaTecla('enter')
def pressEsquerda():
    enviaTecla('left')
def pressDireita():
    enviaTecla('right')
def pressCima():
    enviaTecla('up')
def pressBaixo():
    enviaTecla('down')
def pressInsert():
    enviaTecla('insert')
def pressEsc():
    enviaTecla('esc')
def pressF2():
    enviaTecla('f2')