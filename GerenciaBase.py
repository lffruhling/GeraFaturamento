import funcoes as f
from datetime import datetime

#E-mail Erro
import sys
import traceback
from send_mail import enviar_email_erro

def excecao_handler(excecao_tipo, excecao_valor, trace_back):
    # Captura informações sobre a exceção
    excecao = f"{excecao_tipo.__name__}: {excecao_valor}"
    traceback_info = traceback.format_tb(trace_back)

    # Envia e-mail com informações da exceção
    enviar_email_erro(f"{excecao}\n\nDetalhes do Traceback:\n{''.join(traceback_info)}", 'COrretor Monetário')

# Configura o manipulador de exceções global
sys.excepthook = excecao_handler

def delete(id: int, tabela: str ='faturamento_percentual_cooperativa'):
    vSql = f'DELETE FROM {tabela} WHERE id = %s'
    params = [id]

    conexao = f.conexao()
    cursor = conexao.cursor()

    cursor.execute(vSql, params)
    conexao.commit()

    cursor.close()
    conexao.close()

def retornaDiasFaturamento():
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT dias FROM parametros_faturamento WHERE id = %s', [1])

    result = cursor.fetchone()
    cursor.close()
    conexao.close()

    return int(result[0])
def cooperativas():
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT nome FROM faturamento_percentual_cooperativa ORDER BY nome', [])

    result = cursor.fetchall()
    cursor.close()
    conexao.close()

    return result

def registraFaturamento(cooperativa, dataIni, dataFin, extra):
    vSql = 'SELECT id FROM faturamento_importacao WHERE cooperativa = %s AND data_importacao = %s AND extrajudicial = %s'
    vParams = [cooperativa, datetime.now().strftime('%Y-%m-%d'),extra]
    conexao = f.conexao()
    cursor = conexao.cursor()

    cursor.execute(vSql, vParams)
    result = cursor.fetchone()

    if result is not None:
        vId = result[0]

        vSql = 'SELECT id FROM faturamento_titulos WHERE id_importacao = %s'
        vParams = [vId]
        cursor.execute(vSql, vParams)
        result = cursor.fetchone()

        if result is not None:
            vIdFatTitulos = result[0]

            vSql = 'DELETE FROM faturamento_parcelas WHERE fatura_titulo_id = %s'
            vParams = [vIdFatTitulos]
            cursor.execute(vSql, vParams)
            conexao.commit()

            vSql = 'DELETE FROM faturamento_titulos WHERE id_importacao = %s'
            vParams = [vId]
            cursor.execute(vSql, vParams)
            conexao.commit()

        return vId
    else:
        vSql = 'INSERT INTO faturamento_importacao (cooperativa, data_importacao, inicio_vigencia, final_vigencia, extrajudicial) VALUES (%s, %s, %s, %s, %s)'
        vParams = [cooperativa, datetime.now().strftime('%Y-%m-%d'), dataIni, dataFin, extra]
        cursor.execute(vSql, vParams)
        conexao.commit()

        id = cursor.lastrowid

    cursor.close()
    conexao.close()

    return id

def insereTitulo(cooperativa, titulo, associado, idImportacao):
    db = f.conexao()
    cursor = db.cursor()

    vTitulo = str(titulo).strip()
    agencia = f.identificaAgenciaSicredi(vTitulo)

    cursor.execute('SELECT id FROM faturamento_titulos WHERE titulo_contrato=%s AND cooperativa=%s AND id_importacao = %s;',
                   [vTitulo, cooperativa, idImportacao])
    result = cursor.fetchone()

    if result is None:
        cursor.execute(
            'INSERT INTO faturamento_titulos (titulo_contrato, id_importacao, cooperativa, agencia, associado, data_processamento) VALUE (%s,%s,%s,%s,%s,now())',
            [vTitulo, idImportacao, cooperativa, agencia, associado.rstrip()])
        fTituloId = db.insert_id()

    else:
        cursor.execute(
            "UPDATE faturamento_titulos SET titulo_contrato=%s, cooperativa=%s, agencia=%s, associado=%s, data_processamento=%s  WHERE id=%s",
            [vTitulo, cooperativa, agencia, associado.rstrip(), datetime.now(), result[0]])
        fTituloId = result[0]

        # Limpa tabela de parcelas
        cursor.execute('DELETE FROM faturamento_parcelas WHERE fatura_titulo_id=%s', [fTituloId])

    db.commit()
    cursor.close()
    db.close()

    return fTituloId

def retornaCoop(vConsulta, extra):
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT * FROM faturamento_percentual_cooperativa WHERE cooperativa LIKE %s AND extrajudicial= %s', [f'%{vConsulta}%', extra])
    result = cursor.fetchone()
    cursor.close()
    conexao.close()

    return result


