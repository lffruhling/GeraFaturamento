import os

from _ctypes import sizeof

import constantes
import funcoes as f
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

def abreFicha(arquivo, isTXT=False):
    with open(arquivo, 'r') as reader:
        if not isTXT:
            ficha_grafica = reader.readlines()
            cooperativa = identificaCooperativa(ficha_grafica)
            importaFicha(ficha_grafica, cooperativa)
            # ficha_grafica.close()
        else:
            cooperativa = identificaCooperativa(reader)
            importaFicha(reader, cooperativa)

        return cooperativa

def moeda(valor):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=None)
    return ('R$ %s' % valor)
def identificaCooperativa(ficha):
    ## Identifica a Versão do Arquivo(Se é emitido pelo Sicredi ou da Cresol)
    vlinha = 1

    for linha in ficha:
        if (vlinha <= 20):
            if ("INVEST RAIZES" in linha):
                return constantes.SICREDI_RAIZES
            elif ("CRESOL RAIZ" in linha):
                return constantes.CRESOL_RAIZ
            elif ("INVESTIMENTO CONEXAO" in linha):
                return constantes.SICREDI_CONEXAO
            elif ("CRESOL GERAÇÕES" in linha):
                return constantes.CRESOL_GERACAO

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

def importaFicha(ficha, cooperativa):
    global vFinalVigencia
    global vInicioVigencia

    if cooperativa == constantes.SICREDI_RAIZES or cooperativa == constantes.SICREDI_CONEXAO:
        for linha in ficha:
            if ("TITULO") in linha:
                vTitulo = linha[122:135]
                break

        for linha in ficha:
            if ("ASSOCIADO") in linha:
                vAssociado = linha[16:57]
                break

        fTituloId = insereTitulo(cooperativa, vTitulo, vAssociado)

        db = f.conexao()
        cursor = db.cursor()
        for linha in ficha:
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
    elif cooperativa == constantes.CRESOL_RAIZ or cooperativa == constantes.CRESOL_GERACAO:
        for linha in ficha:
            if ("Nome:") in linha:
                vAssociado = linha[6:len(linha)]
                break

        for linha in ficha:
            if ("Contrato:") in linha:
                vTitulo = linha[10:40]
                break

        fTituloId = insereTitulo(cooperativa, vTitulo, vAssociado)

        db = f.conexao()
        cursor = db.cursor()

        for linha in ficha:

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
        titulos.append({'nro_titulo': titulo[1], "associado": titulo[2], "data_processamento": titulo[3], "parcelas":vParcelas, "total_valor_parcela":moeda(vTotalValorParcelas), "total_faturado": moeda(vTotalFaturado)})
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


    context = {
        "inicio_vigencia": vInicioVigencia,
        "final_vigencia": vFinalVigencia,
        "cooperativa": 'Cresol Raízes',
        "titulos": titulos,
        "total_parcelas": moeda(vTotalParcelas),
        "total_faturamento": moeda(vTotalFatura)
    }
    template = DocxTemplate('C:\\Temp\\Faturamento\\Template.docx')
    template.render(context)
    vDataHora = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    vData =datetime.now().strftime('%m_%Y')
    vNomeArquivo = f'fatura_{vDataHora}'
    vPathArquivo = f'{vPath}/faturamento/{vData}\\'
    utils_f.pastaExiste(f'{vPathArquivo}', True)
    arquivoDoc = f"{vPathArquivo}{vNomeArquivo}.docx"
    template.save(arquivoDoc)
    convert(arquivoDoc, f"{vPathArquivo}{vNomeArquivo}.pdf")
    os.remove(arquivoDoc)

def layout():
    global vInicioVigencia
    global vFinalVigencia
    global vPercentualFaturamento

    sg.theme('reddit')  # Add a touch of color
    # All the stuff inside your window.
    layout = [[sg.T('Fichas', size=10), sg.In(key='pathFichas',disabled=True), sg.FolderBrowse(button_text='Buscar', target='pathFichas')],
              [sg.T('Vigência Inicial', size=17), sg.T('Vigência Final', size=18), sg.T('Percentual de Faturamento', size=25)],
              [sg.InputText(key='dtIni', size=20), sg.In(key='dtFin', size=20),sg.In(key='percentualFat', size=25)],
              [sg.Button(button_text='Cancelar',key='btnCancelar', size=12), sg.Button(button_text='Gerar Fatura', key='btnGeraFatura' ,size=12)]]

    # Create the Window
    window = sg.Window('Gerar Fechamento', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'btnCancelar':  # if user closes window or clicks cancel
            break

        vPercentualFaturamento = values['percentualFat']

        if(len(str(values['pathFichas']).strip())):
            arquivos = os.listdir(values['pathFichas'])
            vInicioVigencia = datetime.strptime(values['dtIni'], '%d/%m/%Y')
            vFinalVigencia = datetime.strptime(values['dtFin'], '%d/%m/%Y')

            for arquivo in arquivos:
                #Ignora diretório de arquivos processados
                if ('faturamento' in arquivo) or ('processados' in arquivo):
                    continue
                vTipoArquivo = arquivo.split(".")[-1]
                if vTipoArquivo.upper() == 'PRN':
                    cooperativa = abreFicha(f"{values['pathFichas']}/{arquivo}")
                    moveFicha(values['pathFichas'], arquivo, cooperativa)
                elif vTipoArquivo.upper() == 'PDF':
                    utils_f.converterPDF(values['pathFichas'], arquivo)
                    cooperativa = abreFicha(f"{values['pathFichas']}/"+str(arquivo).lower().replace("pdf", 'txt'), True)
                    moveFicha(values['pathFichas'], arquivo, cooperativa)  # Move PDF
                    moveFicha(values['pathFichas'], str(arquivo).lower().replace("pdf", 'txt'), cooperativa)  # Move TXT
        else:
            print('Sem arquivos para processar!')

        geraRelatorio(values['pathFichas'])


    window.close()

layout()
