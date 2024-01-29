import PySimpleGUI as sg
import funcoes as f

#E-mail Erro
import sys
import traceback
from send_mail import enviar_email_erro

def excecao_handler(excecao_tipo, excecao_valor, trace_back):
    # Captura informações sobre a exceção
    excecao = f"{excecao_tipo.__name__}: {excecao_valor}"
    traceback_info = traceback.format_tb(trace_back)

    # Envia e-mail com informações da exceção
    enviar_email_erro(f"{excecao}\n\nDetalhes do Traceback:\n{''.join(traceback_info)}", 'Gera Faturamento')

# Configura o manipulador de exceções global
sys.excepthook = excecao_handler

def abre_tela(id=None):
    titulo = 'Nova Taxa'
    cooperativa = None
    nome_apresentacao = None
    percentual = None
    extrajudicial = False
    email = None
    if id is not None:
        titulo = 'Editar Taxa'
        vSql = 'SELECT * FROM faturamento_percentual_cooperativa WHERE id=%s'
        vParams = [id]

        conexao = f.conexao()
        cursor = conexao.cursor()
        cursor.execute(vSql, vParams)
        result = cursor.fetchone()
        conexao.commit()
        cursor.close()
        conexao.close()

        cooperativa         = result[1]
        nome_apresentacao   = result[2]
        percentual          = result[3]
        extrajudicial       = result[4]
        email               = result[5]

        print("carregou os campos?")

    layout = [
        [sg.T('Cooperativa', s=15), sg.I(key="I-cooperativa", s=25, default_text=cooperativa)],
        [sg.T('Nome Apresentação', s=15), sg.I(key="I-nome", s=25, default_text=nome_apresentacao)],
        [sg.T('Percentual', s=15), sg.I(key="I-percentual", s=25, default_text=percentual)],
        [sg.T('Email', s=15), sg.I(key="I-e-mail", s=25, default_text=email)],
        [sg.Check(key="I-extra", text='Extrajudicial', default=extrajudicial, expand_x=True)],
        [sg.HSeparator()],
        [sg.B('Voltar', key='B-VOLTAR', s=15), sg.B('Salvar', key='B-SALVAR', s=15)]
    ]

    window = sg.Window(titulo, layout, modal=True, resizable=True, element_justification='center')

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'B-VOLTAR':
            break
        elif event == "B-SALVAR":
            sql = 'INSERT INTO faturamento_percentual_cooperativa (cooperativa, nome, valor, extrajudicial, email_cooperativa) VALUE (%s,%s,%s,%s,%s)'
            print(values["I-extra"])
            params = [values["I-cooperativa"], values["I-nome"], values["I-percentual"], values["I-extra"],
                      values["I-e-mail"]]

            if id is not None:
                sql = 'UPDATE faturamento_percentual_cooperativa SET cooperativa=%s, nome=%s, valor=%s, extrajudicial=%s, email_cooperativa=%s WHERE id=%s'
                params = [values["I-cooperativa"], values["I-nome"], values["I-percentual"], values["I-extra"],
                          values["I-e-mail"], id]

            conexao = f.conexao()
            cursor = conexao.cursor()
            cursor.execute(sql, params)
            conexao.commit()
            cursor.close()
            conexao.close()
            # sg.popup_no_titlebar('Taxa Adiciona com Sucesso!') COnverter para Notificações do windows
            break

    window.close()