import os
from dotenv import load_dotenv

load_dotenv()

home_dir = os.path.expanduser('~')
afinz_dir = os.path.join(home_dir, 'SOROCRED – CREDITO, FINANCIAMENTO E INVESTIMENTO S')
esteira_dir = os.path.join(afinz_dir, 'Esteira de Integração - Documentos')

CONFIG = {
    'URL_LOGIN': 'https://app.contraktor.com.br/contratos',
    'EMAIL': os.getenv('CONTRAKTOR_EMAIL', 'default@example.com'),
    'PASSWORD': os.getenv('CONTRAKTOR_PASSWORD', 'default_password'),
    'DOWNLOAD_FOLDER': os.path.join(home_dir, 'Downloads'),
    'EXCEL_ESTEIRA': os.path.join(esteira_dir, 'ESTEIRA.xlsx'),
    'EXCEL_PLUXXE': os.path.join(home_dir, 'Downloads', 'PLANSIP4C.xlsx'),
    'LOG_FILE': 'erros_contratos.txt',
    'DEFAULT_TIMEOUT': 60,
    'PDF_TIMEOUT': 10,
    'ATTEMPTS': 3
}
