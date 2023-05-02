import os
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

def abreFicha(pNome):
    with open(vLocalProcessar + pNome, 'r') as reader:
        ficha_grafica = reader.readlines()
        versao = identificaVersaoFicha(ficha_grafica)
        importaFicha(ficha_grafica, versao)
        # ficha_grafica.close()
        return versao

def moeda(valor):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=None)
    return ('R$ %s' % valor)
def identificaVersaoFicha(ficha):
    ## Identifica a Versão do Arquivo(Se é emitido pelo Sicredi ou da Cresol)
    vlinha = 1

    for linha in ficha:
        if (vlinha == 5):
            if ("RAIZES" in linha):
                versao = 'raizes'
            else:
                versao = 'cresol'

            return versao
        vlinha = vlinha + 1

def importaFicha(ficha, versao):
    if versao == 'raizes':
        for linha in ficha:
            if ("TITULO") in linha:
                vTitulo = linha[122:135]
                break

        for linha in ficha:
            if ("ASSOCIADO") in linha:
                vAssociado = linha[16:57]
                break

        db = f.conexao()
        cursor = db.cursor()

        vTitulo = str(vTitulo).strip()

        cursor.execute('SELECT id FROM edersondallabr.fatura_titulos where titulo=%s AND versao=%s;', [vTitulo, versao])
        result = cursor.fetchone()

        if result is None:
            cursor.execute(
                'INSERT INTO fatura_titulos (titulo, versao, associado, data_processamento) VALUE (%s,%s,%s, now())',
                [vTitulo, versao, vAssociado.rstrip()])
            fTituloId = db.insert_id()

        else:
            cursor.execute(
                "UPDATE fatura_titulos SET titulo=%s, versao=%s, associado=%s, data_processamento=%s  WHERE id=%s",
                [vTitulo, versao, vAssociado.rstrip(), datetime.now(), result[0]])
            fTituloId = result[0]

            #Limpa tabela de parcelas
            cursor.execute('DELETE FROM fatura_parcelas WHERE fatura_titulo_id=%s', [fTituloId])

        db.commit()

        vFinalVigencia = datetime.today()
        vInicioVigencia = datetime.today() - timedelta(days=30)
        for linha in ficha:
            if(len(str(linha[0:10]).split("/")) == 3):
                vDataParcela    = linha[0:10]
                dtDataParcela = datetime.strptime(vDataParcela, "%d/%m/%Y")

                if (dtDataParcela > vInicioVigencia) and (dtDataParcela > vFinalVigencia):
                    continue

                vCod            = linha[12:15]
                vHistorico      = linha[17:59]
                if ("AMORTIZACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE TITULO") in vHistorico:
                    vParcela        = linha[59:63]
                    vValor          = linha[90:106]

                    cursor.execute("INSERT INTO fatura_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                                   [fTituloId, dtDataParcela, vCod, vHistorico.rstrip(), vParcela.rstrip(), vValor.lstrip()])

                    print(db.insert_id())
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
                titulo, 
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


arquivos = os.listdir(vLocalProcessar)

if(len(arquivos) > 0):
    for arquivo in arquivos:
        versao = abreFicha(arquivo)
        moveFicha(arquivo, versao)

    geraRelatorio()
else:
    print('Sem arquivos para processar!')