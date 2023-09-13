import funcoes as f

def delete(id: int, tabela: str ='faturamento_percentual_cooperativa'):
    vSql = f'DELETE FROM {tabela} WHERE id = %s'
    params = [id]

    conexao = f.conexao()
    cursor = conexao.cursor()

    cursor.execute(vSql, params)
    conexao.commit()

    cursor.close()
    conexao.close()

