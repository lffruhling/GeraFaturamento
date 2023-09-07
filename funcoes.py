import json
import MySQLdb
import constantes as c
import PySimpleGUI as sg
import os

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

def atualizacaoDisponivel():
    sg.theme('Reddit')

    layout_atualizacao = [    
                [sg.Text(text='Atenção: Existe uma Atualização Disponível', text_color="BLACK", font=("Arial"))],
                [sg.Button('Atualizar Agora', key="btn_atualizar"), sg.Button('Depois', key="btn_cancelar")]      
             ]
    tela_atualizacao = sg.Window('Atualização Disponível', layout_atualizacao, modal=True)

    while True:                    
        eventos, valores = tela_atualizacao.read(timeout=0.1)
        
        if eventos == 'btn_atualizar':            
            if os.path.isfile('C:/Temp/atualizador/atualizador_faturamento.exe'):
                os.startfile('C:/Temp/atualizador/atualizador_faturamento.exe')
                tela_atualizacao.Close()

        if eventos == sg.WINDOW_CLOSED:
            break
        if eventos == 'btn_cancelar':
            tela_atualizacao.close()   

def BuscaUltimaVersao():
    try:
        db = conexao()
        cursor = db.cursor()
        cursor.execute("SELECT versao FROM versao_faturamento order by id desc limit 1")
        resultado = cursor.fetchall()

        if len(resultado) > 0:
            valor = str(resultado[0][0])
            return valor
        else:            
            return None
    except Exception as erro:
        print('Ocorreu um erro ao tentar carregar a última versão gerada. '  + str(erro), True)   

#dados = carregaIndice('igpm', 2022, 3)

