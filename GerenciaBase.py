import funcoes as f
from datetime import datetime

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

def insereTitulo(cooperativa, titulo, associado):
    db = f.conexao()
    cursor = db.cursor()

    vTitulo = str(titulo).strip()
    agencia = f.identificaAgenciaSicredi(vTitulo)

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

def retornaCoop(vConsulta):
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT * FROM faturamento_percentual_cooperativa WHERE cooperativa LIKE %s', [f'%{vConsulta}%'])
    result = cursor.fetchone()
    cursor.close()
    conexao.close()

    return result


