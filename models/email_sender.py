import win32com.client as win32
import pythoncom, os
from config import CONFIG

class EmailSender:
    """Classe para envio de e-mail acumulado da semana"""

    @staticmethod
    def envio_email(anexo):
        try:
            pythoncom.CoInitialize()

            # Verifica se o arquivo existe
            if not os.path.isfile(anexo):
                raise FileNotFoundError(f"Arquivo não encontrado: {anexo}")

            anexo = os.path.abspath(anexo)  # Usa caminho absoluto

            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = 'joao.silva@afinz.com.br'
            mail.Subject = 'Consultores cadastrados'
            mail.Body = "Está é uma lista de todos os consultores cadastrados.\nAtenciosamente,"
            mail.Attachments.Add(anexo)

            mail.Send()
            print("✈️Email enviado com sucesso!")

        except Exception as e:
            print(f"Houve um erro para enviar: {e}")

        finally:
            pythoncom.CoUninitialize()
