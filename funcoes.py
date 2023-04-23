import json
import MySQLdb
import constantes as c

def conexao():
    return MySQLdb.connect(host=c.HOST_DB, user=c.USUARIO_DB, passwd=c.SENHA_DB,db=c.NOME_DB)

def carregaIndice(tabela, ano, mes):
    p_mes = ''
    
    if mes == 1:
        p_mes = 'jan'
    if mes == 2:
        p_mes = 'fev'
    if mes == 3:
        p_mes = 'mar'
    if mes == 4:
        p_mes = 'abr'
    if mes == 5:
        p_mes = 'mai'
    if mes == 6:
        p_mes = 'jun'
    if mes == 7:
        p_mes = 'jul'
    if mes == 8:
        p_mes = 'ago'
    if mes == 9:
        p_mes = 'set'
    if mes == 10:
        p_mes = 'out'
    if mes == 11:
        p_mes = 'nov'
    if mes == 12:
        p_mes = 'dez'

    try:
        db = conexao()
        cursor = db.cursor()
        cursor.execute("SELECT " + p_mes + " FROM " + tabela + " WHERE ano = " + str(ano))
        resultado = cursor.fetchall()
    except Exception as erro:
        print('Ocorreu um erro ao tentar carregar o indice ' + str(tabela) + '. ' + str(erro))

    print(resultado[0][0])
    if len(resultado) > 0:
        valor = str(resultado[0][0]*100)
        return round(float(valor),2)
    else:
        print('Sem resultados')
        return 0

#dados = carregaIndice('igpm', 2022, 3)

