import telas.config_dias as TLDIAS
import telas.parametros.index as TLCONFIG
import GerenciaBase as base
import funcoes as f
import GerenciaImportacoes as imp
import GerenciaRelatorios as rel
import util.funcoes as utils_f

from datetime import datetime, timedelta

import PySimpleGUI as sg

def atualizaBarraProgresso(tela, texto=None, vMax=None, vAtual=None, corTexto = 'Blue'):
    if texto is not None:
        tela['-LABEL_PROGRESS-'].Update(texto, text_color=corTexto)

    if vAtual is not None:
        tela['-PROGRESS_BAR-'].Update(max=vMax, current_count=vAtual)
        if vMax is not None:
            tela['-PROGRESS_BAR-'].Update(max=vMax)

    tela.refresh()

def abre_tela():
    nomeCoop = None
    versaoExe = '1.0.2'
    verificaVersao = True

    dias = base.retornaDiasFaturamento()

    now = datetime.now()

    vDataIniF = (now - timedelta(days=dias)).strftime('%d/%m/%Y')
    vDataFinF = now.strftime('%d/%m/%Y')

    # ------ Thema Layout ------ #
    sg.theme('reddit')

    # ------ Menu ------ #

    menu_def = [['Configurações', ["Taxas Cooperativas", "Dias de Intervalo"]]]

    # ------ Layout Tela Principal------ #
    layout = [
        [sg.MenubarCustom(menu_def, tearoff=False, bar_background_color="#EEEEEE", bar_text_color="#000000",background_color="#EEEEEE", text_color="#000000")],
        [
            sg.Text('                  '),
            sg.Image(filename='logo1.png'),
            sg.Text(text='Gera Faturamento', text_color="Black", font=("Arial", 22, "bold"), expand_x=True,justification='left')
        ],
        [sg.HSeparator()],
        [
            sg.T('Fichas', s=11),
            sg.I(key='I-arquivos', justification="l", disabled=True, enable_events=True, change_submits=True),
            sg.FilesBrowse(button_text='Buscar', s=12, target='I-arquivos')
        ],
        [
            sg.T('Vigência Inicial', s=11),
            sg.I(vDataIniF, key='dtIni', s=11, justification="l"),
            sg.T('Vigência Final', s=11),
            sg.I(vDataFinF, key='dtFin', s=11, justification="l")
        ],
        [
            sg.T('Cooperativa:', s=11),
            sg.Combo(values=['Padrão'], key="C-cooperativas", enable_events=True, expand_x=True),
            sg.Check(key="I-extra", text='Extrajudicial', default=False, expand_x=True)
        ],
        [sg.HSeparator()],
        [
            sg.B(button_text='Cancelar', key='btnCancelar', s=12),
            sg.B(button_text='Gerar Fatura', key='btnGeraFatura', s=12),
            sg.Text('Versão: ' + versaoExe, expand_x=True, justification='right'),
            sg.T(text='', key="-LABEL_PROGRESS-")
        ],
        [sg.ProgressBar(100, orientation='h', key='-PROGRESS_BAR-', bar_color=("#338AFF", "#D1D7DF"), expand_x=True, size=(40, 5))],
    ]
    cooperativas = []
    # Create the Window
    sg.set_options(dpi_awareness=True)
    window = sg.Window('Gerar Fechamento', layout, resizable=True, grab_anywhere=True)
    vPercentual = None
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read(timeout=0.1)

        if len(cooperativas) == 0:
            cooperativas.append('Padrão')
            coops = base.cooperativas()
            for coop in coops:
                cooperativas.append(coop[0])
            window['C-cooperativas'].Update(values=cooperativas)
            window['C-cooperativas'].Update('Padrão')


        if event == sg.WIN_CLOSED or event == 'btnCancelar':  # if user closes window or clicks cancel
            break

        if event == 'Taxas Cooperativas':
            # tela_config_taxas()
            TLCONFIG.abre_tela()
            continue

        if event == "Dias de Intervalo":
            TLDIAS.abre_tela()
            dias = base.retornaDiasFaturamento()
            now = datetime.now()

            vDataIniF = (now - timedelta(days=dias)).strftime('%d/%m/%Y')
            window['dtIni'].Update(value=vDataIniF)
            continue

        if verificaVersao:
            verificaVersao = False
            versaoBanco = f.BuscaUltimaVersao()
            if versaoExe != versaoBanco:
                f.atualizacaoDisponivel()

        if event == 'btnGeraFatura':
            vInicioVigencia = utils_f.formata_data_banco(values['dtIni'], '%d/%m/%Y', '%Y-%m-%d')
            vFinalVigencia = utils_f.formata_data_banco(values['dtFin'], '%d/%m/%Y', '%Y-%m-%d')

            vInicioVigenciaCalc = datetime.strptime(values['dtIni'], '%d/%m/%Y')
            vFinalVigenciaCalc = datetime.strptime(values['dtFin'], '%d/%m/%Y')

            id = base.registraFaturamento(nomeCoop, vInicioVigencia, vFinalVigencia, values['I-extra'])

            vArquivos = values['I-arquivos']
            vListaArquivos = vArquivos.split(';')
            if (len(vListaArquivos) > 0):
                for arquivo in vListaArquivos:
                    vTipoArquivo = arquivo.split(".")[-1]
                    if vTipoArquivo.upper() == 'PDF':
                        if not utils_f.arquivoExiste(str(arquivo).lower().replace("pdf", 'txt')):
                            utils_f.converterPDF(arquivo)

                        arquivo = str(arquivo).lower().replace("pdf", 'txt')

                    nomeCoop, vPercentual = f.identificaCooperativaCombo(window, arquivo,values['I-extra'])
                    imp.importaFicha(arquivo, sg, window, id, vInicioVigenciaCalc, vFinalVigenciaCalc, values['I-extra'], nomeCoop, vPercentual)
            else:
                sg.popup_no_titlebar('Sem arquivos para processar!')

            rel.geraRelatorio("c:/TEMP", id, values['dtIni'], values['dtFin'], vPercentual, nomeCoop, values['I-extra'])
        elif event == 'I-arquivos':
            atualizaBarraProgresso(window, texto='Preparando sistema para importação! Aguarde...')
            vArquivos = values['I-arquivos']
            vListaArquivos = vArquivos.split(';')
            arquivo = vListaArquivos[0]
            nomeCoop, vPercentual = f.identificaCooperativaCombo(window, arquivo, values['I-extra'])

    window.close()
