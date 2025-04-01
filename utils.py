import logging
import time
import glob
import os
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from config import CONFIG

def setup_logging():
    """Configura o sistema de logging."""
    logging.basicConfig(
        filename=CONFIG['LOG_FILE'],
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    # Adicionar também log no console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    logging.getLogger('').addHandler(console)
    
def abreviar_nome(nome_completo):
    """
    Abrevia o nome, mantendo o primeiro e último nomes completos.
    
    Args:
        nome_completo (str): Nome completo da pessoa
        
    Returns:
        str: Nome abreviado
    """
    partes = nome_completo.split()
    if len(partes) <= 2:
        return nome_completo
        
    primeiro_nome = partes[0]
    ultimo_nome = partes[-1]
    nomes_abreviados = [f"{parte[0]}." for parte in partes[1:-1]]
    return f"{primeiro_nome} {' '.join(nomes_abreviados)} {ultimo_nome}"

# Dicionário de estados para UFs
ESTADOS_UF = {
    "Acre": "AC", "Alagoas": "AL", "Amapá": "AP", "Amazonas": "AM", "Bahia": "BA", "Ceará": "CE",
    "Distrito Federal": "DF", "Espírito Santo": "ES", "Goiás": "GO", "Maranhão": "MA", "Mato Grosso": "MT",
    "Mato Grosso do Sul": "MS", "Minas Gerais": "MG", "Pará": "PA", "Paraíba": "PB", "Paraná": "PR",
    "Pernambuco": "PE", "Piauí": "PI", "Rio de Janeiro": "RJ", "Rio Grande do Norte": "RN", "Rio Grande do Sul": "RS",
    "Rondônia": "RO", "Roraima": "RR", "Santa Catarina": "SC", "São Paulo": "SP", "Sergipe": "SE", "Tocantins": "TO"
}

def estado_para_uf(estado):
    """Converte nome do estado para UF."""
    return ESTADOS_UF.get(estado, estado)

def aguardar_elemento(driver, by, valor, tipo_espera='presenca', tempo=None):
    """
    Aguarda que um elemento esteja disponível na página.
    
    Args:
        driver: WebDriver da Selenium
        by: Tipo de localizador (By.ID, By.XPATH, etc)
        valor: Valor do localizador
        tipo_espera: Tipo de espera ('presenca', 'clicavel', 'invisibilidade')
        tempo: Tempo máximo de espera em segundos
        
    Returns:
        WebElement se encontrado, None caso contrário
    """
    if tempo is None:
        tempo = CONFIG['DEFAULT_TIMEOUT']
        
    try:
        if tipo_espera == 'presenca':
            return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((by, valor)))
        elif tipo_espera == 'clicavel':
            return WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((by, valor)))
        elif tipo_espera == 'invisibilidade':
            return WebDriverWait(driver, tempo).until(EC.invisibility_of_element_located((by, valor)))
    except TimeoutException:
        logging.error(f"Timeout ao esperar pelo elemento: {valor}")
        return None

def encontrar_arquivo_recente(pasta, padrao="*.pdf", timeout=None):
    """
    Encontra o arquivo mais recente que corresponde ao padrão na pasta.
    
    Args:
        pasta: Caminho da pasta onde procurar
        padrao: Padrão de arquivo a procurar (glob)
        timeout: Tempo máximo de espera em segundos
        
    Returns:
        Path do arquivo encontrado ou None
    """
    if timeout is None:
        timeout = CONFIG['PDF_TIMEOUT']
        
    logging.info(f"Procurando arquivo com padrão '{padrao}' na pasta {pasta}")
    tempo_inicio = time.time()

    while time.time() - tempo_inicio < timeout:
        caminho_padrao = os.path.join(pasta, padrao)
        arquivos_encontrados = glob.glob(caminho_padrao)

        if arquivos_encontrados:
            arquivo_mais_recente = max(arquivos_encontrados, key=os.path.getmtime)
            # Verifica se o arquivo está completo (não está sendo escrito)
            tamanho_inicial = os.path.getsize(arquivo_mais_recente)
            time.sleep(1)
            if os.path.getsize(arquivo_mais_recente) == tamanho_inicial:
                logging.info(f"Arquivo encontrado: {arquivo_mais_recente}")
                return arquivo_mais_recente
        time.sleep(2)

    logging.warning(f"Timeout: Nenhum arquivo encontrado após {timeout} segundos")
    return None

def excluir_arquivo(arquivo):
    """Remove um arquivo do sistema."""
    try:
        os.remove(arquivo)
        logging.info(f"Arquivo {arquivo} excluído com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao excluir o arquivo {arquivo}: {e}")