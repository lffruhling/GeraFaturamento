import funcoes as f
import PySimpleGUI as sg

def abre_tela():
    conexao = f.conexao()
    cursor = conexao.cursor()
    cursor.execute('SELECT dias FROM parametros_faturamento WHERE id = %s', [1])

    result = cursor.fetchone()
    cursor.close()
    conexao.close()
    dias = result[0]

    layout = [
        [sg.T('Dias de Intervalo'), sg.I(default_text=dias, key='I_DIAS')],
        [sg.B('Salvar', key='B_SAVE', s=10), sg.B('Voltar', key='B_VOLTAR', s=10)]
    ]

    window = sg.Window('Dias de Intervalos', layout, modal=True)

    while True:
        event, values = window.read(timeout=0.1)

        if event == sg.WIN_CLOSED or event == 'B_VOLTAR':  # if user closes window or clicks cancel
            break

        if event == 'B_SAVE':
            print(values['I_DIAS'])
            conexao = f.conexao()
            cursor = conexao.cursor()
            cursor.execute('UPDATE parametros_faturamento SET dias = %s WHERE id = %s', [values['I_DIAS'], 1])
            conexao.commit()
            cursor.close()
            conexao.close()
            sg.popup_ok('Dias alterados com sucesso!')
            break

    window.close()