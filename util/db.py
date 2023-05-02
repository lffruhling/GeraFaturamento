import util.criptografia as criptografia
import MySQLdb
import os
#from dotenv import load_dotenv
    
try:
  load_dotenv()   
except:
  print('Falha ao tentar carregar o arquivo .env')
    
#Essas informações adicionar em um arquivo .Env e adicioná-lo ao .gitignore
def conectaDB():
  try:  
    MySQLdb.connect(host='10.4.21.24', database='credenciais', user=os.environ["DB_CREDENCIAIS_USUARIO"], password=os.environ["DB_CREDENCIAIS_SENHA"])   
  except Exception as erro:
    print('Falha ao tentar conectar ao bando de dados das credenciais. ' + str(erro))
    
def consultaDB1(expressao):
  try:
    
    conexao = conectaDB()
    cursor = conexao.cursor()
    
    cursor.execute(expressao)
    resultado = cursor.fetchall()
    
    cursor.close()
    conexao.close()
    
    return resultado
        
  except Exception as erro:
    print('Falha ao tentar executar consulta ao DB. ' + str(erro))
    
def insertSQL(vsql,vparametros):
  try:  
    conexao = conectaDB()
    cursor = conexao.cursor()  
    
    cursor.execute(vsql,vparametros)    
    conexao.commit()
    
    cursor.close()
    conexao.close()
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar executar a instrução SQL. ' + str(erro))
    
def updateSQL(vsql):
  try:  
    conexao = conectaDB()
    cursor = conexao.cursor()  
    
    cursor.execute(vsql)    
    conexao.commit()
    
    cursor.close()
    conexao.close()
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar executar a instrução SQL. ' + str(erro))
    
def recuperaCredencial(vnome):
  try:
    print('Executa consulta de chave no DB')  
    
    conexao = conectaDB()
    cursor = conexao.cursor()  
    
    vparam = [(vnome.upper())]
    cursor.execute("SELECT usuario, senha FROM credenciais where upper(credenciais.nome) = %s", vparam)
    resultado = cursor.fetchall()
    
    usuario = str(resultado[0][0])
    senha   = criptografia.descriptografar(str(resultado[0][1]))
    
    vresultado = {
      'usuario' : usuario,
      'senha'   : senha
    }
    
    return vresultado
    
  except Exception as erro:
    print('Ocorreu um erro ao tentar consultar dados no DB. ' + str(erro))
  