import os
import funcoes as f
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import locale
from docx2pdf import convert

vLocal = f'C:\\Temp\\Faturamento\\Processar\\04_2023\\'
template = DocxTemplate('C:\\Users\\lffru\\PycharmProjects\\GeraFaturas\\relatorios\\Template.docx')
def abreFicha(pNome):
    with open(vLocal + pNome, 'r') as reader:
        ficha_grafica = reader.readlines()
        versao = identificaVersaoFicha(ficha_grafica)
        print(versao)
        importaFicha(ficha_grafica, versao)
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
        vlinha = 1

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
        cursor.execute(
            'INSERT INTO fatura_titulos (titulo, versao, associado, data_processamento) VALUE (%s,%s,%s, now())',
            [vTitulo, versao, vAssociado.rstrip()])


        fTituloId = db.insert_id()
        db.commit()

        print(fTituloId)
        vFinalVigencia = datetime.today()
        vInicioVigencia = datetime.today() - timedelta(days=30)
        for linha in ficha:
            if(len(str(linha[0:10]).split("/")) == 3):
                vDataParcela    = linha[0:10]
                dtDataParcela = datetime.strptime(vDataParcela, "%d/%m/%Y")
                print("d1 is greater than d2 : ", dtDataParcela < vInicioVigencia)
                print("d1 is less than d2 : ", dtDataParcela < vFinalVigencia)

                if (dtDataParcela > vInicioVigencia) and (dtDataParcela > vFinalVigencia):
                    continue

                vCod            = linha[12:15]
                vHistorico      = linha[17:59]
                if ("AMORTIZACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE PARCELA") in vHistorico or ("LIQUIDACAO DE TITULO") in vHistorico:
                    vParcela        = linha[59:63]
                    vValor          = linha[90:106]

                    cursor.execute("INSERT INTO fatura_detalhe (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                                   [fTituloId, dtDataParcela, vCod, vHistorico.rstrip(), vParcela.rstrip(), vValor.lstrip()])

                    print(db.insert_id())
                    db.commit()
        cursor.close()
        db.close()
            # if (("AMORTIZACAO DE PARCELA" in linha) or ("LIQUIDACAO DE PARCELA" in linha) or (
            #         "LIQUIDACAO DE TITULO" in linha)):
            #     valor = linha[82:107].replace(" ", "")
            #     amortizacoes.append(valor.replace(".", ""))
            # vlinha = vlinha + 1
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
                FROM fatura_detalhe where fatura_titulo_id = %s
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
            FROM fatura_detalhe AS fd 
	            INNER JOIN edersondallabr.fatura_titulos AS ft
		            ON fd.fatura_titulo_id = ft.id
            """
    cursor.execute(sql)
    rTotalParcelas = cursor.fetchone()
    vTotalParcelas = rTotalParcelas[0]
    vPercentualCobranca = 10
    vTotalFatura = ((vTotalParcelas * vPercentualCobranca)/100) + vTotalParcelas
    print([vTotalParcelas, vPercentualCobranca, vTotalFatura])
    print(moeda(vTotalFatura))

    context = {
        "titulos": titulos,
        "total_parcelas": moeda(vTotalParcelas),
        "total_faturamento": moeda(vTotalFatura)
    }

    template.render(context)
    template.save('C:\\Users\\lffru\\PycharmProjects\\GeraFaturas\\relatorios\\Fatura-ok.docx')
    convert("C:\\Users\\lffru\\PycharmProjects\\GeraFaturas\\relatorios\\Fatura-ok.docx",
            "C:\\Users\\lffru\\PycharmProjects\\GeraFaturas\\relatorios\\Fatura-ok.pdf")

arquivos = os.listdir(vLocal)

for arquivo in arquivos:
    abreFicha(arquivo)

geraRelatorio()