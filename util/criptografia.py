#import rsa
#from rsa import key
import datetime
from datetime import datetime
from datetime import date
import util.db as db
import os

#from cryptography.fernet import Fernet

def log(msg):
    print(datetime.now().strftime("%m/%d/%Y %H:%M:%S") + ' => ' + msg + "\n")
    
vchave = []

def chaveCriptografia():
    try:
        log('Verificando chave de criptografia das credenciais.')
        resultado = db.consultaChave()  
        if not resultado:
            log('Não existe chave de criptografia criada. Retorno => ' + str(resultado))
            log('Gerando nova chave.')
            
            ##Caso nao exista chave no DB, cria uma nova chave, recupera ela e armazena no DB
            (pubkey, privkey) = rsa.newkeys(128)

            try:
                db.criaChave('geral', privkey, pubkey)
                
                vchave.append(pubkey)
                vchave.append(privkey)
                
                log('Chaves de criptografia criadas com sucesso!')
            except Exception as erro:
                log('Falha ao tentar gerar a chave de Criptografia. ' + str(erro))                                        
        else:
            log('Chave recuperada com sucesso. ')
            pubkey  = str(resultado[0][1])
            privkey = str(resultado[0][0])

            vchave.append(pubkey)
            vchave.append(privkey)
            
        return vchave
                
    except Exception as erro:
        log('Falha ao tentar gerar/consultar a chave de Criptografia. ' + str(erro))        
        
def criptografar(vtexto):
    mensagem     = vtexto.encode()
    criptografia = Fernet(recuperaCriptografia())
    
    retorno = criptografia.encrypt(mensagem) 
    
    return retorno.decode()

def descriptografar(vtexto):
    criptografia = Fernet(recuperaCriptografia())
    
    retorno = criptografia.decrypt(vtexto.encode())
    return retorno.decode()

def recuperaCriptografia():
    vlocal = 'F:\\PROCESSOS E INOVAÇÃO\\Python\\RSA-KEY\\criptografia.key'
    
    if os.path.isfile(vlocal):        
        print('Recuperando chaves de criptografia...')        
        retorno = open(vlocal, "rb").read()
    else:
        print('Gerando novas chaves de criptografia...')        
        chave = Fernet.generate_key()
        with open(vlocal, "wb") as arquivo_keys:
            arquivo_keys.write(chave)
        
        retorno = open(vlocal, "rb").read()
        
    return retorno        