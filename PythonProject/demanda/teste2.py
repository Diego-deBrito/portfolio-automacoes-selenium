# Função para enviar e-mails com anexos.
import win32com.client as win32
import time


def enviar_email_tecnico(email_destino: str, numero_processo: str, destinatario: str,
                         caminho_pasta: str):
    mensagem = f"""

        <p>Prezado(a) {destinatario},</p>

        <p>Gostaria de informar que a proposta {numero_processo} passou por uma atualização. A versão revisada já está disponível para consulta no TransfereGov.</p>

        <p> Por favor, revise as alterações e me informe caso tenha alguma dúvida.</p>

        <p>O documendo baixado se encontra na pasta{caminho_pasta}, com o nome {numero_processo.replace('/', '_')}</p>

        <p> Atenciosamente,</p>


        <p> Atenciosamente,</p>


    """

    try:
        # Cria integração com o outrlook
        outlook = win32.Dispatch('outlook.application')

        # Configurar e-mail
        email = outlook.CreateItem(0)
        email.To = f'{email_destino}'
        email.Subject = f'Atualização na Proposta {numero_processo}'
        email.HTMLBody = mensagem

        time.sleep(1)
        email.Display()
        time.sleep(10)
        email.Send()
        print(f"✅ E-mail enviado para {destinatario}, no endereço {email_destino}")

    except Exception as e:
        print(f"❌ Falha ao enviar e-mail para {destinatario}: \n{e}")



enviar_email_tecnico('andrei.rodrigues@esporte.gov.br', 'xxxx', 'andrei.rodrigues@esporte.gov.br', 'xxxxx')