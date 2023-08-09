import os
import funcoes as f

import constantes
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import locale
from docx2pdf import convert
import shutil

import util.funcoes as utils_f

import PySimpleGUI as sg

# vLocalProcessar = f'C:\\Temp\\Faturamento\\Processar\\'
# vLocalProcessados = f'C:\\Temp\\Faturamento\\Processado\\'
# vLocalRelatorios = f'C:\\Temp\\Faturamento\\Relatorios\\'
array_datas      = ['01/','02/','03/','04/','05/','06/','07/','08/','09/','10/','11/','12/',
                    '13/','14/','15/','16/','17/','18/','19/','20/','21/','22/','23/','24/',
                    '25/','26/','27/','28/','29/','30/','31/']

#Variáveis Globais
vFinalVigencia          = None
vInicioVigencia         = None
vPercentualFaturamento  = None

def atualizaBanco(values):
    sql = 'UPDATE faturamento_percentual_cooperativa SET cooperativa=%s, nome=%s, valor=%s WHERE id=%s'
    params = [values[1], values[2], values[3], values[0]]

    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute(sql, params)
    conexao.commit()
    cursor.close()
    conexao.close()

def edit_cell(windowTelaConfig, key, row, col, justify='left'):

    global textvariable, edit

    def callback(event, row, col, text, key):
        global edit
        # event.widget gives you the same entry widget we created earlier
        widget = event.widget
        if key == 'Focus_Out':
            # Get new text that has been typed into widget
            text = widget.get()
        # Destroy the entry widget
        widget.destroy()
        # Destroy all widgets
        widget.master.destroy()
        # Get the row from the table that was edited
        # table variable exists here because it was called before the callback
        values = list(table.item(row, 'values'))
        # Store new value in the appropriate row and column
        if(values[col] != text):
            values[col] = text
            table.item(row, values=values)
            atualizaBanco(values)
        edit = False


    if edit or row <= 0 or col <= 0:
        return

    edit = True
    # Get the Tkinter functionality for our window
    root = windowTelaConfig.TKroot
    # Gets the Widget object from the PySimpleGUI table - a PySimpleGUI table is really
    # what's called a TreeView widget in TKinter
    table = windowTelaConfig[key].Widget
    # Get the row as a dict using .item function and get individual value using [col]
    # Get currently selected value
    text = table.item(row, "values")[col]

    # Return x and y position of cell as well as width and height (in TreeView widget)
    x, y, width, height = table.bbox(row, col)

    # Create a new container that acts as container for the editable text input widget
    frame = sg.tk.Frame(root)
    # c
    tamanhoCpoTitulo = windowTelaConfig['-L-tituloConfig'].get_size()
    tamanhoCpoSperador = windowTelaConfig['-S-separadorConfig'].get_size()
    alturaFinalY = y + tamanhoCpoTitulo[1] + tamanhoCpoSperador[1]
    alturaFinalY += 15
    alturaFinalX = x + 3

    frame.place(x=alturaFinalX, y=alturaFinalY, anchor="nw", width=width, height=height)

    # textvariable represents a text value
    textvariable = sg.tk.StringVar()
    textvariable.set(text)
    # Used to acceot single line text input from user - editable text input
    # frame is the parent window, textvariable is the initial value, justify is the position
    entry = sg.tk.Entry(frame, textvariable=textvariable, justify=justify)
    # Organizes widgets into blocks before putting them into the parent
    entry.pack()
    # selects all text in the entry input widget
    entry.select_range(0, sg.tk.END)
    # Puts cursor at end of input text
    entry.icursor(sg.tk.END)
    # Forces focus on the entry widget (actually when the user clicks because this initiates all this Tkinter stuff, e
    # ending with a focus on what has been created)
    entry.focus_force()
    # When you click outside of the selected widget, everything is returned back to normal
    # lambda e generates an empty function, which is turned into an event function
    # which corresponds to the "FocusOut" (clicking outside of the cell) event
    entry.bind("<FocusOut>", lambda e, r=row, c=col, t=text, k='Focus_Out':callback(e, r, c, t, k))

def adicionaCoop(table):
    global dados
    print(dados)
    layout = [
        [sg.T('Cooperativa', s=15), sg.I(key="I-cooperativa", s=25, justification='right')],
        [sg.T('Nome Apresentação', s=15), sg.I(key="I-nome", s=25, justification='right')],
        [sg.T('Percentual', s=15), sg.I(key="I-percentual", s=25, justification='right')],
        [sg.B('Cancelar', key='B-cancelar', s=15), sg.B('Salvar', key='B-salvar', s=15)]
    ]

    window = sg.Window("Nova Taxa", layout, modal=True, resizable=True, element_justification='center')

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'B-cancelar':
            break
        elif event == "B-salvar":
            sql = 'INSERT INTO faturamento_percentual_cooperativa (cooperativa, nome, valor) VALUE (%s,%s,%s)'
            params = [values["I-cooperativa"], values["I-nome"], values["I-percentual"]]

            conexao = f.conexao()
            cursor = conexao.cursor()
            cursor.execute(sql, params)
            conexao.commit()
            cursor.close()
            conexao.close()
            dados.append([cursor.lastrowid, values["I-cooperativa"],values["I-nome"], values["I-percentual"]])
            table.update(dados)
            sg.popup_no_titlebar('Taxa Adiciona com Sucesso!')
            break

    window.close()

def tela_config_taxas():
    global edit
    global dados
    edit = False
    # ------ Layout Tela Taxas ------ #
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT * FROM faturamento_percentual_cooperativa')

    dados = list()
    # image_data = i.btn_delete_img
    for vRow in cursor.fetchall():
        dados.append([vRow[0], vRow[1], vRow[2], vRow[3], 'X'])

    cursor.close()
    conexao.close()


    layout = [
        [sg.T('Cofigurações de Taxas', key="-L-tituloConfig")],
        [sg.HSeparator(key="-S-separadorConfig")],
        [sg.Table(values=dados, headings=['###', 'COOPERATIVA', 'NOME', 'PERCENTUAL', 'REMOVER'],
                  auto_size_columns=True,
                  max_col_width=25,
                  num_rows=10,
                  # alternating_row_color=sg.theme_button_color()[1],
                  key="-TABLE-",
                  row_height=20,
                  tooltip='Tabela de Percentuais',
                  expand_x=True,
                  expand_y=True,
                  # enable_events=True,
                  # bind_return_key=True
                  enable_click_events=True
                  )
         ],
        [sg.B('Adicionar', key='B-NOVA', s=15)]
        # [sg.B('Cancelar', key='btnCancelarTaxas', s=15), sg.B('Salvar', key='btnSalvarTaxas', s=15)]
    ]

    sg.set_options(dpi_awareness=True)
    windowTelaConfig = sg.Window("Configurações de Taxas", layout, modal=True, resizable=True, element_justification='center')

    while True:
        event, values = windowTelaConfig.read()

        if event == sg.WIN_CLOSED or event == 'btnCancelarTaxas':
            break
        elif isinstance(event, tuple):
            table = windowTelaConfig['-TABLE-']
            if isinstance(event[2][0], int) and event[2][0] > -1:
                row, col = event[2]
                if col == (len(table.ColumnHeadings) -1):
                    result = sg.popup_ok_cancel('Deseja Remover esse item?')
                    if result == 'OK':
                        id = dados[row][0]
                        sql = 'DELETE FROM faturamento_percentual_cooperativa WHERE id=%s'
                        params = [id]
                        conexao = f.conexao()
                        cursor = conexao.cursor()
                        cursor.execute(sql,params)
                        conexao.commit()
                        cursor.close()
                        conexao.close()
                        dados.pop(row)
                        table.update(dados)
                        sg.popup_no_titlebar('Registro Removido com sucesso!')
                        continue
            edit_cell(windowTelaConfig, '-TABLE-', row+1, col, justify="right")
        elif event == 'B-NOVA':
            adicionaCoop(windowTelaConfig['-TABLE-'])
            # adicionaCoop()

    windowTelaConfig.close()


def moeda(valor):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=None)
    return ('R$ %s' % valor)
def identificaCooperativa(file):
    ## Identifica a Versão do Arquivo(Se é emitido pelo Sicredi ou da Cresol)
    vlinha = 1
    vConsulta = None

    for linha in file:
        if (vlinha <= 20):
            if ("INVEST RAIZES" in linha):
                vConsulta = "INVEST RAIZES"
            elif ("CRESOL RAIZ" in linha):
                vConsulta = "CRESOL RAIZ"
            elif ("INVESTIMENTO CONEXAO" in linha):
                vConsulta = "INVESTIMENTO CONEXAO"
            elif ("CRESOL GERAÇÕES" in linha):
                vConsulta = "CRESOL GERAÇÕES"

            if vConsulta is not None:
                conexao = f.conexao()
                cursor = conexao.cursor()
                cursor.execute('SELECT * FROM faturamento_percentual_cooperativa WHERE cooperativa LIKE %s', [f'%{vConsulta}%'])
                result = cursor.fetchone()
                cursor.close()
                conexao.close()

                if result is None:
                    return None

                return [result[2], result[3]]
                break
        vlinha = vlinha + 1


def identificaAgenciaSicredi(titulo):
    return titulo[2:4]
def insereTitulo(cooperativa, titulo, associado):
    db = f.conexao()
    cursor = db.cursor()

    vTitulo = str(titulo).strip()
    agencia = identificaAgenciaSicredi(vTitulo)

    cursor.execute('SELECT id FROM fatura_titulos where titulo_contrato=%s AND cooperativa=%s;',
                   [vTitulo, cooperativa])
    result = cursor.fetchone()

    if result is None:
        cursor.execute(
            'INSERT INTO fatura_titulos (titulo_contrato, cooperativa, agencia, associado, data_processamento) VALUE (%s,%s,%s,%s,now())',
            [vTitulo, cooperativa, agencia, associado.rstrip()])
        fTituloId = db.insert_id()

    else:
        cursor.execute(
            "UPDATE fatura_titulos SET titulo_contrato=%s, cooperativa=%s, agencia=%s, associado=%s, data_processamento=%s  WHERE id=%s",
            [vTitulo, cooperativa, agencia, associado.rstrip(), datetime.now(), result[0]])
        fTituloId = result[0]

        # Limpa tabela de parcelas
        cursor.execute('DELETE FROM fatura_parcelas WHERE fatura_titulo_id=%s', [fTituloId])

    db.commit()
    cursor.close()
    db.close()

    return fTituloId

def importaFicha(arquivo, isTXT=False):
    global vFinalVigencia
    global vInicioVigencia
    global vPercentualFaturamento

    with open(arquivo, 'r') as reader:
        if not isTXT:
            ficha_grafica = reader.readlines()
        else:
            ficha_grafica = reader

        cooperativa, vPercentualFaturamento = identificaCooperativa(ficha_grafica)

        if cooperativa is None:
            sg.popup_no_titlebar('Cooperativa Não Localizada! Processamento será abortado')
            raise "Cooperativa Não Encontrada"

        if 'SICREDI' in str(cooperativa).upper():
            for linha in ficha_grafica:
                if ("TITULO") in linha:
                    vTitulo = linha[122:135]
                    break

            for linha in ficha_grafica:
                if ("ASSOCIADO") in linha:
                    vAssociado = linha[16:57]
                    break

            fTituloId = insereTitulo(cooperativa, vTitulo, vAssociado)

            db = f.conexao()
            cursor = db.cursor()
            for linha in ficha_grafica:
                if(len(str(linha[0:10]).split("/")) == 3):
                    vDataParcela    = linha[0:10]
                    dtDataParcela = datetime.strptime(vDataParcela, "%d/%m/%Y")
                    if (dtDataParcela >= vInicioVigencia) and (dtDataParcela >= vFinalVigencia):
                        #Caso uma delas esteja fora do intervalo não deixa adicionar
                        continue
                    elif not (dtDataParcela > vInicioVigencia) and not (dtDataParcela > vFinalVigencia):
                        #Caso as Duas datas seja Falsas, No caso as duas estão fora do intervalo
                        continue

                    vCod            = linha[12:15]
                    vHistorico      = linha[17:59]
                    if ("AMORTIZACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE TITULO") in vHistorico:
                        vParcela        = linha[59:63]
                        vValor          = linha[90:106]

                        cursor.execute("INSERT INTO fatura_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                                       [fTituloId, dtDataParcela, vCod, vHistorico.rstrip(), vParcela.rstrip(), vValor.lstrip()])

                        # print(db.insert_id())
            db.commit()
            cursor.close()
            db.close()
        elif 'CRESOL' in str(cooperativa).upper():
            for linha in ficha_grafica:
                if ("Nome:") in linha:
                    vAssociado = linha[6:len(linha)]
                    break

            for linha in ficha_grafica:
                if ("Contrato:") in linha:
                    vTitulo = linha[10:40]
                    break

            fTituloId = insereTitulo(cooperativa, vTitulo, vAssociado)

            db = f.conexao()
            cursor = db.cursor()
            for linha in ficha_grafica:

                ## Divide a linha em um array de 4 partes
                # linha_atual = linha.split(" ", 4)
                linha_atual = linha.split(" ")

                vLinhaLancamento = False
                ## identifica se é uma linha de lançamento de movimentação
                if (linha_atual[1][0:3] in array_datas):
                    vLinhaLancamento = True
                    ## Se for linha de lançamento, captura as informações

                if vLinhaLancamento:
                    descricao = ''
                    parcela = int(linha_atual[0])

                    data_m = linha_atual[2].split('/')
                    dia_m = data_m[0]
                    mes_m = data_m[1]
                    ano_m = data_m[2]
                    data_movimento = datetime(int(ano_m), int(mes_m), int(dia_m))

                    operacao = int(linha_atual[3])

                    descricao = ''
                    textoLista = linha_atual[4:len(linha_atual) - 4]
                    for texto in textoLista:
                        descricao = f'{descricao} {texto}'
                    print(descricao)

                    ## Captura valor, substitui virgulas por ponto, converte em float
                    str_valor = linha_atual[len(linha_atual) - 4].replace('.', '')
                    str_valor = str_valor.replace(',', '.')
                    valor = float(str_valor)

                    if (data_movimento >= vInicioVigencia) and (data_movimento >= vFinalVigencia):
                        #Caso uma delas esteja fora do intervalo não deixa adicionar
                        continue
                    elif not (data_movimento > vInicioVigencia) and not (data_movimento > vFinalVigencia):
                        #Caso as Duas datas seja Falsas, No caso as duas estão fora do intervalo
                        continue

                    print(data_movimento)
                    if ("AMORTIZAÇÃO") in descricao or ("LIQUIDACAO DE PARCELA") in descricao or ("LIQUIDACAO DE TITULO") in descricao:

                        cursor.execute(
                            "INSERT INTO fatura_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                            [fTituloId, data_movimento, operacao, descricao, parcela, valor])
            db.commit()
            cursor.close()
            db.close()
def moveFicha(vPath, vNomeArquivo, cooperativa):
    vVigencia = datetime.now().strftime('%m_%Y')
    vMoverPara = f'{vPath}\\processados\\{vVigencia}\\{cooperativa}'
    utils_f.pastaExiste(vMoverPara, True)
    shutil.move(f"{vPath}\\{vNomeArquivo}", f'{vMoverPara}\\{vNomeArquivo}')

def geraRelatorio(vPath):
    global vInicioVigencia
    global vFinalVigencia
    global vPercentualFaturamento

    # To-DO
    # Se fro Sicredi, agrupar por cooperativa e por agencia para falicitar o rateio!

    print(vPercentualFaturamento)

    db = f.conexao()
    cursor = db.cursor()
    sql = """
            SELECT 
                id,
                titulo_contrato, 
                associado, 
                data_processamento,
                cooperativa,
                agencia
            FROM fatura_titulos ORDER BY cooperativa, agencia, associado;
		"""
    cursor.execute(sql)
    rTitulos = cursor.fetchall()

    titulos = []
    for titulo in rTitulos:
        sql = """
                SELECT 
                    data_parcela, 
                    historico,
                    valor,
                    parcela 
                FROM fatura_parcelas where fatura_titulo_id = %s
        """
        cursor.execute(sql, [titulo[0]])
        rParcelas = cursor.fetchall()
        vParcelas = []
        vTotalValorParcelas = 0
        for parcela in rParcelas:
            vParcelas.append({"data": parcela[0].strftime("%d/%m/%Y"), "historico":parcela[1], "valor":moeda(parcela[2]),"parcela":parcela[3], "valor_faturado":moeda(parcela[2] / 10)})
            vTotalValorParcelas += parcela[2]
        if len(rParcelas) == 0:
            vParcelas.append({"data": "--", "historico": "Sem Lancamentos para este Título", "valor": "--"})
        vTotalFaturado = (vTotalValorParcelas * float(vPercentualFaturamento)) / 100
        titulos.append({'nro_titulo': titulo[1], "associado": titulo[2], "data_processamento": titulo[3], "cooperativa": titulo[4], "agencia": titulo[5], "parcelas":vParcelas, "total_valor_parcela":moeda(vTotalValorParcelas), "total_faturado": moeda(vTotalFaturado)})
    sql = """
            SELECT 
	            coalesce(sum(valor), 0) AS total_parcelas
            FROM fatura_parcelas AS fd 
	            INNER JOIN edersondallabr.fatura_titulos AS ft
		            ON fd.fatura_titulo_id = ft.id
            """
    cursor.execute(sql)
    rTotalParcelas = cursor.fetchone()
    vTotalParcelas = rTotalParcelas[0]
    if vTotalParcelas > 0:
        vTotalFatura = ((vTotalParcelas * float(vPercentualFaturamento))/100)
    else:
        vTotalFatura = 0

    sqlRateio = """
                    SELECT 
                        ft.cooperativa,
                        ft.agencia,
                        coalesce(sum(fp.valor),0) as valor
                    FROM fatura_titulos as ft
                        LEFT JOIN fatura_parcelas as fp
                            ON ft.id = fp.fatura_titulo_id
                    GROUP BY fp.fatura_titulo_id, ft.cooperativa
                    ORDER BY ft.cooperativa, ft.agencia;
                """
    cursor.execute(sqlRateio)
    rRateios = cursor.fetchall()

    rateios = []
    for rateio in rRateios:
        rateios.append({'cooperativa': rateio[0], "agencia": rateio[1], "valor": moeda(rateio[2])})

    context = {
        "inicio_vigencia": vInicioVigencia,
        "final_vigencia": vFinalVigencia,
        "cooperativa": 'Cresol Raízes',
        "titulos": titulos,
        "rateios": rateios,
        "total_parcelas": moeda(vTotalParcelas),
        "total_faturamento": moeda(vTotalFatura)
    }
    template = DocxTemplate('C:\\Temp\\Faturamento\\Template.docx')
    template.render(context)
    vDataHora = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    vData =datetime.now().strftime('%m_%Y')
    vNomeArquivo = f'fatura_{vDataHora}'
    if vPath == '':
        vPath = 'c:\\Temp'
        utils_f.pastaExiste(vPath, True)
    vPathArquivo = f'{vPath}/faturamento/{vData}\\'
    utils_f.pastaExiste(f'{vPathArquivo}', True)
    arquivoDoc = f"{vPathArquivo}{vNomeArquivo}.docx"
    template.save(arquivoDoc)
    convert(arquivoDoc, f"{vPathArquivo}{vNomeArquivo}.pdf")
    os.remove(arquivoDoc)
    result = sg.popup_ok('Faturamento Gerado com Sucesso!')
    if result == 'OK':
        os.startfile(vPathArquivo)
        os.startfile(f"{vPathArquivo}{vNomeArquivo}.pdf")

def layout():
    global vInicioVigencia
    global vFinalVigencia
    global vPercentualFaturamento

    now = datetime.now()
    vDataIniF = now.strftime('%d/%m/%Y')
    vDataFinF = now.strftime('%d/%m/%Y')

    # ------ Thema Layout ------ #
    sg.theme('reddit')

    # ------ Menu ------ #

    menu_def = [['Configurações', ["Taxas Cooperativas", "Dias de Intervalo"]]]

    # ------ Layout Tela Principal------ #
    layout = [
        [sg.MenubarCustom(menu_def, tearoff=False, bar_background_color="#EEEEEE", bar_text_color="#000000", background_color="#EEEEEE", text_color="#000000")],
        [sg.HSeparator()],
        [sg.T('Fichas', s=11), sg.I(key='I-arquivos', justification="l", disabled=True), sg.FilesBrowse(button_text='Buscar', s=12, target='I-arquivos')],
        [sg.T('Vigência Inicial', s=11), sg.I(vDataIniF, key='dtIni', s=11, justification="l"), sg.T('Vigência Final', s=11), sg.I(vDataFinF,key='dtFin', s=11, justification="l")],
        [sg.HSeparator()],
        [sg.B(button_text='Cancelar', key='btnCancelar', s=12), sg.B(button_text='Gerar Fatura', key='btnGeraFatura', s=12)]
    ]

    # Create the Window
    sg.set_options(dpi_awareness=True)
    window = sg.Window('Gerar Fechamento', layout, resizable=True, grab_anywhere=True)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'btnCancelar':  # if user closes window or clicks cancel
            break

        if event == 'Taxas Cooperativas':
            tela_config_taxas()
            continue

        # vPercentualFaturamento = values['percentualFat']
        vInicioVigencia = datetime.strptime(values['dtIni'], '%d/%m/%Y')
        vFinalVigencia = datetime.strptime(values['dtFin'], '%d/%m/%Y')

        vArquivos = values['I-arquivos']
        vListaArquivos = vArquivos.split(';')
        if(len(vListaArquivos)>0):
            for arquivo in vListaArquivos:

                vTipoArquivo = arquivo.split(".")[-1]
                if vTipoArquivo.upper() == 'PRN':
                    importaFicha(arquivo)
                elif vTipoArquivo.upper() == 'PDF':
                    utils_f.converterPDF(values['pathFichas'], arquivo)
                    importaFicha(str(arquivo).lower().replace("pdf", 'txt'), True)
        else:
            sg.popup_no_titlebar('Sem arquivos para processar!')

        geraRelatorio("c:/TEMP")


    window.close()

layout()