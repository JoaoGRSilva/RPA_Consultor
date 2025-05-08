import win32com.client as win32
import os
from config import CONFIG

class EmailSender:
    """Classe para envio de e-mail acumulado da semana"""

    @staticmethod
    def envio_email():

        attachment = mail.Attachments.Add(CONFIG['COMPILADO_FOLDER'])
        if not os.path.exists(attachment):
            raise FileNotFoundError(f"Arquivo não encontrado: {attachment}")
        
        try:
            # Abertura do Outlook
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            # Configurar o destinatário
            mail.To = 'talita.souza@afinz.com.br'
            
            # Configurar o assunto
            mail.Subject = 'Teste'
            
            # Adicionar algum corpo ao email 
            mail.Body = "Este é um email de teste."
            
            # Adicionar o anexo
            mail.Attachments.Add(attachment)
            
            # Enviar o email
            mail.Send()
        except Exception as e:
            print(f"Houve um erro para enviar: {e}")