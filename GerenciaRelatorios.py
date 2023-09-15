import funcoes as f
from docxtpl import DocxTemplate
from datetime import datetime
import os
from docx2pdf import convert
import util.funcoes as utils_f

def geraRelatorio(vPath, idImportacao, vInicioVigencia, vFinalVigencia, vPercentualFaturamento):
    # To-DO
    # Se fro Sicredi, agrupar por cooperativa e por agencia para falicitar o rateio!

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
            FROM faturamento_titulos
            WHERE id_importacao = %s 
            ORDER BY cooperativa, agencia, associado;
		"""
    cursor.execute(sql, [idImportacao])
    rTitulos = cursor.fetchall()

    titulos = []
    for titulo in rTitulos:
        sql = """
                SELECT 
                    data_parcela, 
                    historico,
                    valor,
                    parcela 
                FROM faturamento_parcelas 
                WHERE fatura_titulo_id = %s
        """
        cursor.execute(sql, [titulo[0]])
        rParcelas = cursor.fetchall()
        vParcelas = []
        vTotalValorParcelas = 0
        for parcela in rParcelas:
            vParcelas.append({"data": parcela[0].strftime("%d/%m/%Y"), "historico":parcela[1], "valor":f.moeda(parcela[2]),"parcela":parcela[3], "valor_faturado":f.moeda(parcela[2] / 10)})
            vTotalValorParcelas += parcela[2]
        if len(rParcelas) == 0:
            vParcelas.append({"data": "--", "historico": "Sem Lancamentos para este Título", "valor": "--"})
        vTotalFaturado = (vTotalValorParcelas * float(vPercentualFaturamento)) / 100
        titulos.append({'nro_titulo': titulo[1], "associado": titulo[2], "data_processamento": titulo[3], "cooperativa": titulo[4], "agencia": titulo[5], "parcelas":vParcelas, "total_valor_parcela":f.moeda(vTotalValorParcelas), "total_faturado": f.moeda(vTotalFaturado)})
    sql = """
            SELECT 
	            coalesce(sum(valor), 0) AS total_parcelas
            FROM faturamento_parcelas AS fd 
	            INNER JOIN faturamento_titulos AS ft
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
                    FROM faturamento_titulos as ft
                        LEFT JOIN faturamento_parcelas as fp
                            ON ft.id = fp.fatura_titulo_id
                    GROUP BY fp.fatura_titulo_id, ft.cooperativa
                    ORDER BY ft.cooperativa, ft.agencia;
                """
    cursor.execute(sqlRateio)
    rRateios = cursor.fetchall()

    rateios = []
    for rateio in rRateios:
        rateios.append({'cooperativa': rateio[0], "agencia": rateio[1], "valor": f.moeda(rateio[2])})

    context = {
        "inicio_vigencia": vInicioVigencia,
        "final_vigencia": vFinalVigencia,
        "cooperativa": 'Cresol Raízes',
        "titulos": titulos,
        "rateios": rateios,
        "total_parcelas": f.moeda(vTotalParcelas),
        "total_faturamento": f.moeda(vTotalFatura)
    }
    template = DocxTemplate('Template.docx')
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
    # result = sg.popup_ok('Faturamento Gerado com Sucesso!')
    # if result == 'OK':
    os.startfile(vPathArquivo)
    os.startfile(f"{vPathArquivo}{vNomeArquivo}.pdf")