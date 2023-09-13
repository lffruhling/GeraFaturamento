import PySimpleGUI as sg
import funcoes as f
import GerenciaBase as base
import telas.parametros.c_e_campos as cps
from datetime import datetime

def atualizaTabela(tela):
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('''
        SELECT 
    	    id, 
            cooperativa, 
            nome, 
            CONCAT(valor, "%") as valor, 
            CASE WHEN extrajudicial THEN 'Sim' ELSE 'Não' END as extrajudicial, 
            COALESCE(email_cooperativa, '') as email_cooperativa 
        FROM 
            faturamento_percentual_cooperativa
        ''')
    dados = cursor.fetchall()
    cursor.close()
    conexao.close()

    tela['-TABLE-'].Update(values=dados)

    return dados
def abre_tela():

    layout = [
        [sg.T('Cofigurações de Taxas', key="-L-tituloConfig")],
        [sg.HSeparator(key="-S-separadorConfig")],
        [sg.Table(values=[], headings=['###', 'COOPERATIVA', 'NOME', ' % ', 'EXTRA', 'E-MAIL COOP.'],
                  max_col_width = 35,
                  auto_size_columns=True,
                  num_rows=10,
                  key="-TABLE-",
                  row_height=20,
                  tooltip='Tabela de Percentuais',
                  expand_x=True,
                  expand_y=True,
                  enable_events=True
                  )
         ],
        [
            sg.B('Voltar', key='B-VOLTAR', s=15),
            sg.B('Adicionar', key='B-NOVA', s=15),
            sg.B('Editar', key='B-EDITAR', s=15,  disabled=True),
            sg.B('Remover', key='B-REMOVER', s=15,  disabled=True)
        ]
    ]

    id = None
    window = sg.Window("Configurações de Taxas", layout, modal=True, resizable=True, element_justification='center')

    carregaDados = False
    hora_atualizacao = datetime.now()

    while True:
        event, values = window.read(timeout=0.1)

        atualizar = hora_atualizacao - datetime.now()
        segundos = atualizar.total_seconds()

        if (segundos * -1) >= 30:
            carregaDados = False

        if not (carregaDados):
            hora_atualizacao = datetime.now()
            dados = atualizaTabela(window)
            carregaDados = True

        if event == sg.WIN_CLOSED or event == 'B-VOLTAR':
            break
        elif event == 'B-NOVA':
            cps.abre_tela()
            hora_atualizacao = datetime.now()
            dados = atualizaTabela(window)
        elif event == 'B-EDITAR':
            cps.abre_tela(id)
            hora_atualizacao = datetime.now()
            dados = atualizaTabela(window)
        elif event == 'B-REMOVER':
            result = sg.popup_ok_cancel('Deseja realmente remover esse registro?')

            if result == 'OK':
                base.delete(id)
                hora_atualizacao = datetime.now()
                dados = atualizaTabela(window)

        elif event == "-TABLE-":
            if len(values[event]) > 0:
                id = dados[values[event][0]][0]
                window['B-EDITAR'].Update(disabled=False)
                window['B-REMOVER'].Update(disabled=False)
    window.close()