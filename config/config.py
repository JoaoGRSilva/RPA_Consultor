import sys
import os
import glob
from dotenv import load_dotenv

load_dotenv()

# Diretório do executável (ou script quando em desenvolvimento)
exec_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
home_dir = os.path.expanduser('~')
afinz_dir = os.path.join(home_dir, 'SOROCRED – CREDITO, FINANCIAMENTO E INVESTIMENTO S')
esteira_dir = os.path.join(afinz_dir, 'Esteira de Integração - Documentos')
rpa_dir = os.path.join(esteira_dir, '#RPA')

# Encontrar qualquer arquivo .xlsx no diretório do executável
xlsx_files = glob.glob(os.path.join(exec_dir, '*.xlsx'))
excel_pluxxe = xlsx_files[0] if xlsx_files else 'pluxxe.xlsx'  

CONFIG = {
    'URL_LOGIN': 'https://app.contraktor.com.br/contratos',
    'EMAIL': os.getenv('CONTRAKTOR_EMAIL', 'default@example.com'),
    'PASSWORD': os.getenv('CONTRAKTOR_PASSWORD', 'default_password'),
    'DOWNLOAD_FOLDER': os.path.join(home_dir, 'Downloads'),
    'EXCEL_ESTEIRA': os.path.join(esteira_dir, 'ESTEIRA.xlsx'),
    'EXCEL_PLUXXE': excel_pluxxe,
    'EXCEL_CONTRATOS': os.path.join(rpa_dir, '#RPA.xlsx'),
    'LOG_FILE': os.path.join(exec_dir, 'erros_contratos.txt'), 
    'DEFAULT_TIMEOUT': 60,
    'PDF_TIMEOUT': 10,
    'ATTEMPTS': 3
}