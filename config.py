import os
from dotenv import load_dotenv

load_dotenv()


CONFIG = {
    'URL_LOGIN': 'https://app.contraktor.com.br/contratos',
    'EMAIL': os.getenv('CONTRAKTOR_EMAIL', 'default@example.com'),
    'PASSWORD': os.getenv('CONTRAKTOR_PASSWORD', 'default_password'),
    'DOWNLOAD_FOLDER': os.path.join(os.path.expanduser('~'), 'Downloads'),
    'EXCEL_ESTEIRA': os.path.join(os.path.expanduser('~'), 'Downloads', 'ESTEIRA.xlsx'),
    'EXCEL_PLUXXE': os.path.join(os.path.expanduser('~'), 'Downloads', 'PLANSIP4C_3230687_110325.xlsx'),
    'LOG_FILE': 'erros_contratos.txt',
    'DEFAULT_TIMEOUT': 60,
    'PDF_TIMEOUT': 10,
}
