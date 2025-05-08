import win32com.client as win32
import pythoncom
from config import CONFIG

class EmailSender:
    """Classe para envio de e-mail acumulado da semana"""

    @staticmethod
    def envio_email():
        """
        attachment = mail.Attachments.Add(CONFIG['COMPILADO_FOLDER'])
        if not os.path.exists(attachment):
            raise FileNotFoundError(f"Arquivo não encontrado: {attachment}")    
        """
        
        try:
            pythoncom.CoInitialize()
            
            # Abertura do Outlook
            outlook = win32.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
    
            mail = outlook.CreateItem(0)
            if mail is None:
                raise Exception("Não foi possível criar o item de email")
                
            # Configurar o destinatário
            mail.To = 'carla.romao@afinz.com.br'
            
            # Configurar o assunto
            mail.Subject = 'Teste'
            
            # Adicionar algum corpo ao email 
            mail.Body = "Este é um email de teste."
            
            # mail.Attachments.Add(attachment)
            
            mail.Save()
            mail.Send()
            
            print("Email enviado com sucesso!")
            
        except Exception as e:
            print(f"Houve um erro para enviar: {e}")
            
        finally:
            pythoncom.CoUninitialize()