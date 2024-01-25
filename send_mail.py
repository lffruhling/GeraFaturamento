import smtplib
import traceback

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def enviar_email_erro(excecao, nome_aplicacao):
    username = 'leonardofruhling@gmail.com'
    password = 'runm ttop mxzp ksxs'
    mail_from = 'leonardofruhling@gmail.com'
    mail_to = 'leonardofruhling@gmail.com'
    mail_subject = f"Erro na Aplicação {nome_aplicacao}"
    mail_body = f"""\
    O seguinte erro ocorreu nna Aplicação {nome_aplicacao}:

    {excecao}

    Detalhes do Traceback:
    {traceback.format_exc()}
    """

    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject
    mimemsg.attach(MIMEText(mail_body, 'plain'))
    try:
        connection = smtplib.SMTP(host='smtp.gmail.com', port=587)
        connection.starttls()
        connection.login(username,password)
        connection.send_message(mimemsg)
        # connection.quit()
        print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Erro ao enviar e-mail: {e}')
    finally:
        # Fechar a conexão com o servidor SMTP
        connection.quit()