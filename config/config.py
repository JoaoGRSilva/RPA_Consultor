import sys
import os
from dotenv import load_dotenv

load_dotenv()

# Diretório do executável (ou script quando em desenvolvimento)
exec_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
home_dir = os.path.expanduser('~')
afinz_dir = os.path.join(home_dir, 'SOROCRED – CREDITO, FINANCIAMENTO E INVESTIMENTO S')
esteira_dir = os.path.join(afinz_dir, 'Esteira de Integração - Documentos')
modelo_dir = os.path.join(esteira_dir, 'MODELO')
rpa_dir = os.path.join(esteira_dir, '#RPA')


CONFIG = {
    'URL_LOGIN': 'https://app.contraktor.com.br/contratos',
    'EMAIL': os.getenv('CONTRAKTOR_EMAIL', 'default@example.com'),
    'PASSWORD': os.getenv('CONTRAKTOR_PASSWORD', 'default_password'),
    'DOWNLOAD_FOLDER': os.path.join(home_dir, 'Downloads'),
    'EXCEL_ESTEIRA': os.path.join(esteira_dir, 'ESTEIRA.xlsx'),
    'EXCEL_PLUXXE': os.path.join(modelo_dir, 'PLANSIP4C_MODELO.xlsx'),
    'EXCEL_CONTRATOS': os.path.join(rpa_dir, '#RPA.xlsx'),
    'PLUXXE_FOLDER': os.path.join(esteira_dir, 'PLUXXE'),
    'COMPILADO_FOLDER': os.path.join(esteira_dir, 'COMPILADO'),
    'LOG_FILE': os.path.join(exec_dir, 'erros_contratos.txt'), 
    'DEFAULT_TIMEOUT': 60,
    'PDF_TIMEOUT': 10,
    'ATTEMPTS': 3
}