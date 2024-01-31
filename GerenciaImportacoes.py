import funcoes as f
import GerenciaBase as base
from datetime import datetime
import util.funcoes as utils_f

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

array_datas      = ['01/','02/','03/','04/','05/','06/','07/','08/','09/','10/','11/','12/',
                    '13/','14/','15/','16/','17/','18/','19/','20/','21/','22/','23/','24/',
                    '25/','26/','27/','28/','29/','30/','31/']

def importaFicha(arquivo, sg, tela, idImportacao, vFinalVigencia, vInicioVigencia, extra, cooperativa, vPercentualFaturamento):
    isTXT = str(arquivo.split(".")[-1]).lower() == 'txt'

    # with open(arquivo, 'r') as ficha_grafica:
    #     # cooperativa, vPercentualFaturamento = f.identificaCooperativa(ficha_grafica, tela, extra)
    #     if cooperativa is None:
    #         sg.popup_no_titlebar('Cooperativa Não Localizada! Processamento será abortado')
    #         ficha_grafica.close()
    #         raise "Cooperativa Não Encontrada"
    #     ficha_grafica.close()
    #
    #
    with open(arquivo, 'r') as ficha_grafica:

        if 'SICREDI' in str(cooperativa).upper():
            for linha in ficha_grafica:
                if ("TITULO") in linha:
                    if isTXT:
                        vTitulo = linha[109:120]
                    else:
                        vTitulo = linha[122:135]
                    break

            for linha in ficha_grafica:
                if ("ASSOCIADO") in linha:
                    if isTXT:
                        vAssociado = linha[16:45]
                    else:
                        vAssociado = linha[16:57]
                    break

            fTituloId = base.insereTitulo(cooperativa, vTitulo, vAssociado, idImportacao)

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

                        cursor.execute("INSERT INTO faturamento_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
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

            fTituloId = base.insereTitulo(cooperativa, vTitulo, vAssociado, idImportacao)

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

                    if ("AMORTIZAÇÃO") in descricao or ("LIQUIDACAO DE PARCELA") in descricao or ("LIQUIDACAO DE TITULO") in descricao:

                        cursor.execute(
                            "INSERT INTO faturamento_parcelas (fatura_titulo_id, data_parcela, cod, historico, parcela, valor) VALUE (%s,%s,%s,%s,%s,%s)",
                            [fTituloId, data_movimento, operacao, descricao, parcela, valor])
            db.commit()
            cursor.close()
            db.close()

    ficha_grafica.close()
    utils_f.removeArquivo(arquivo)