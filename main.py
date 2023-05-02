import os

import constantes
import funcoes as f
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import locale
from docx2pdf import convert
import shutil

import util.funcoes as utils_f

vLocalProcessar = f'C:\\Temp\\Faturamento\\Processar\\'
vLocalProcessados = f'C:\\Temp\\Faturamento\\Processado\\'
vLocalRelatorios = f'C:\\Temp\\Faturamento\\Relatorios\\'

def abreFicha(pNome, isTXT=False):
    with open(vLocalProcessar + pNome, 'r') as reader:
        if not isTXT:
            ficha_grafica = reader.readlines()
            versao = identificaVersaoFicha(ficha_grafica)
            importaFicha(ficha_grafica, versao)
            # ficha_grafica.close()
        else:
            versao = identificaVersaoFicha(reader)
            importaFicha(reader,versao)

        return versao

def moeda(valor):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=None)
    return ('R$ %s' % valor)
def identificaVersaoFicha(ficha):
    ## Identifica a Versão do Arquivo(Se é emitido pelo Sicredi ou da Cresol)
    vlinha = 1

    for linha in ficha:
        if (vlinha <= 20):
            if ("COOP CRED POUP E INVEST RAIZES" in linha):
                return 'sicredi_raizes'
            elif ("CRESOL RAIZ" in linha):
                return 'cresol_raizes'

        vlinha = vlinha + 1

def insereTitulo(versao, titulo, associado):
    db = f.conexao()
    cursor = db.cursor()

    vTitulo = str(titulo).strip()

    cursor.execute('SELECT id FROM fatura_titulos where titulo_contrato=%s AND versao=%s;',
                   [vTitulo, versao])
    result = cursor.fetchone()

    if result is None:
        cursor.execute(
            'INSERT INTO fatura_titulos (titulo_contrato, versao, associado, data_processamento) VALUE (%s,%s,%s, now())',
            [vTitulo, versao, associado.rstrip()])
        fTituloId = db.insert_id()

    else:
        cursor.execute(
            "UPDATE fatura_titulos SET titulo_contrato=%s, versao=%s, associado=%s, data_processamento=%s  WHERE id=%s",
            [vTitulo, versao, associado.rstrip(), datetime.now(), result[0]])
        fTituloId = result[0]

        # Limpa tabela de parcelas
        cursor.execute('DELETE FROM fatura_parcelas WHERE fatura_titulo_id=%s', [fTituloId])

    db.commit()
    cursor.close()
    db.close()

    return fTituloId

def importaFicha(ficha, versao):
    vFinalVigencia = datetime.today()
    vInicioVigencia = datetime.today() - timedelta(days=30)
    if versao == 'sicredi_raizes':
        for linha in ficha:
            if ("TITULO") in linha:
                vTitulo = linha[122:135]
                break

        for linha in ficha:
            if ("ASSOCIADO") in linha:
                vAssociado = linha[16:57]
                break

        fTituloId = insereTitulo(versao, vTitulo, vAssociado)

        db = f.conexao()
        cursor = db.cursor()
        for linha in ficha:
            if(len(str(linha[0:10]).split("/")) == 3):
                vDataParcela    = linha[0:10]
                dtDataParcela = datetime.strptime(vDataParcela, "%d/%m/%Y")
                # print(f'Maior que data de inicio: {(dtDataParcela > vInicioVigencia)}')
                # print(f'Maior que data de Final: {(dtDataParcela > vFinalVigencia)}')
                # print('-------------------------------------------')
                if (dtDataParcela > vInicioVigencia) and (dtDataParcela > vFinalVigencia):
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
    elif versao == 'cresol_raizes':
        for linha in ficha:
            if ("Contrato:") in linha:
                vTitulo = linha[10:40]
                break

        for linha in ficha:
            if ("Nome:") in linha:
                vAssociado = linha[7:len(linha)]
                break
        fTituloId = insereTitulo(versao, vTitulo, vAssociado)

        db = f.conexao()
        cursor = db.cursor()

        for linha in ficha:

            if(len(str(linha[13:25]).split("/")) == 3):
                vParcela = linha[0:2]
                linhaData = linha[13:23]
                linhaCod = linha[24:28]
                linhaHistorico = linha[29:len(linha)]
                linhaValor = linha[50:len(linha)]

                try:
                    iParcela = int(vParcela)
                except Exception:
                    print('Não é parcela')
                    iParcela = 0

                if (iParcela >= 10 and iParcela <= 99):
                    linhaData = linha[14:24]
                    linhaCod = linha[25:29]
                    linhaHistorico = linha[30:len(linha)]
                    linhaValor = linha[51:len(linha)]

                elif (iParcela >= 100):
                    linhaData = linha[15:25]
                    linhaCod = linha[16:26]
                    linhaHistorico = linha[31:len(linha)]
                    linhaValor = linha[52:len(linha)]

                vDataParcela = linhaData

                if ':' in vDataParcela:
                    continue
                print(vDataParcela)
                dtDataParcela = datetime.strptime(vDataParcela, "%d/%m/%Y")
                print(f'Maior que data de inicio: {(dtDataParcela > vInicioVigencia)}')
                print(f'Maior que data de Final: {(dtDataParcela > vFinalVigencia)}')
                print('-------------------------------------------')
                if (dtDataParcela > vInicioVigencia) and (dtDataParcela > vFinalVigencia):
                    #Caso uma delas esteja fora do intervalo não deixa adicionar
                    continue
                elif not (dtDataParcela > vInicioVigencia) and not (dtDataParcela > vFinalVigencia):
                    #Caso as Duas datas seja Falsas, No caso as duas estão fora do intervalo
                    continue

                vCod = linhaCod
                vHistorico = linhaHistorico
                vHistoricoArray = vHistorico.split(",")
                vHistorico = vHistoricoArray[0]
                for char, replacement in constantes.NUMEROS:
                    if char in vHistorico:
                        vHistorico = vHistorico.replace(char, replacement)

                if ("AMORTIZAÇÃO") in vHistorico or ("LIQUIDACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE TITULO") in vHistorico:

                    vValor = linhaValor
                    vValorArray = vValor.split(",")
                    vValor = vValorArray[0] + "," + vValorArray[1]
                    for char, replacement in constantes.ALFABETO:
                        if char in vValor:
                            vValor = vValor.replace(char, replacement)

                    cursor.execute(
                        "INSERT INTO fatura_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                        [fTituloId, dtDataParcela, vCod, vHistorico.rstrip(), vParcela.rstrip(), vValor.lstrip()])
        db.commit()
        cursor.close()
        db.close()
def moveFicha(vNomeArquivo, versao):
    #verificar se Existe Diretório
    #'%m_%Y'
    vVigencia = datetime.now().strftime('%m_%Y')
    vMoverPara = f'{vLocalProcessados}{vVigencia}\\{versao}'
    utils_f.pastaExiste(vMoverPara, True)
    shutil.move(vLocalProcessar + vNomeArquivo, f'{vMoverPara}\\{vNomeArquivo}')

def geraRelatorio():
    db = f.conexao()
    cursor = db.cursor()
    sql = """
            SELECT 
                id,
                titulo_contrato, 
                associado, 
                data_processamento
            FROM fatura_titulos;
		"""
    cursor.execute(sql)
    rTitulos = cursor.fetchall()

    titulos = []
    for titulo in rTitulos:
        sql = """
                SELECT 
                    data_parcela, 
                    historico, 
                    valor 
                FROM fatura_parcelas where fatura_titulo_id = %s
        """
        cursor.execute(sql, [titulo[0]])
        rParcelas = cursor.fetchall()
        vParcelas = []
        for parcela in rParcelas:
            vParcelas.append({"data":parcela[0].strftime("%d/%m/%Y") ,"historico":parcela[1], "valor":moeda(parcela[2])})
        if len(rParcelas) == 0:
            vParcelas.append({"data": "--", "historico": "Sem Lancamentos para este Título", "valor": "--"})
        titulos.append({'nro_titulo': titulo[1], "associado": titulo[2], "data_processamento": titulo[3], "parcelas":vParcelas})
    sql = """
            SELECT 
	            sum(valor) AS total_parcelas
            FROM fatura_parcelas AS fd 
	            INNER JOIN edersondallabr.fatura_titulos AS ft
		            ON fd.fatura_titulo_id = ft.id
            """
    cursor.execute(sql)
    rTotalParcelas = cursor.fetchone()
    vTotalParcelas = rTotalParcelas[0]
    vPercentualCobranca = 10
    vTotalFatura = ((vTotalParcelas * vPercentualCobranca)/100) + vTotalParcelas

    context = {
        "titulos": titulos,
        "total_parcelas": moeda(vTotalParcelas),
        "total_faturamento": moeda(vTotalFatura)
    }
    template = DocxTemplate('C:\\Temp\\Faturamento\\Template.docx')
    template.render(context)
    vDataHora = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    vData =datetime.now().strftime('%m_%Y')
    vNomeArquivo = f'fatura_{vDataHora}'
    vPathArquivo = f'{vLocalRelatorios}{vData}\\'
    utils_f.pastaExiste(f'{vPathArquivo}', True)
    arquivoDoc = f"{vPathArquivo}{vNomeArquivo}.docx"
    template.save(arquivoDoc)
    convert(arquivoDoc, f"{vPathArquivo}{vNomeArquivo}.pdf")
    os.remove(arquivoDoc)

def main():
    arquivos = os.listdir(vLocalProcessar)

    if(len(arquivos) > 0):
        for arquivo in arquivos:
            vTipoArquivo = arquivo.split(".")[-1]
            if vTipoArquivo.upper() == 'PRN':
                versao = abreFicha(arquivo)
                moveFicha(arquivo, versao)
            elif vTipoArquivo.upper() == 'PDF':
                utils_f.converterPDF(vLocalProcessar, arquivo)
                versao = abreFicha(str(arquivo).lower().replace("pdf",'txt'), True)
                moveFicha(arquivo, versao) #Move PDF
                moveFicha(str(arquivo).lower().replace("pdf",'txt'), versao) #Move TXT


        geraRelatorio()
    else:
        print('Sem arquivos para processar!')

main()