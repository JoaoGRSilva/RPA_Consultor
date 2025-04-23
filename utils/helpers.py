import logging
import time
import glob
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from config.config import CONFIG

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
    nomes_abreviados = [parte[0].upper() for parte in partes[1:-1]]
    return f"{primeiro_nome} {' '.join(nomes_abreviados)} {ultimo_nome}"


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
        print(f"Timeout ao esperar pelo elemento: {valor}")
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
        
    arquivos_antes = set(glob.glob(os.path.join(pasta, padrao)))
    tempo_inicio = time.time()

    while time.time() - tempo_inicio < timeout:
        arquivos_agora = set(glob.glob(os.path.join(pasta, padrao)))
        novos_arquivos = list(arquivos_agora - arquivos_antes)

        if novos_arquivos:
            novos_arquivos.sort(key=lambda x: os.path.getctime(x))
            return novos_arquivos[-1]
        time.sleep(1)
    return None

def excluir_arquivo(arquivo):
    """Remove um arquivo do sistema."""
    try:
        os.remove(arquivo)
    except Exception as e:
        print(f"Erro ao excluir o arquivo {arquivo}: {e}")

def try_click(element):
    tentativas = CONFIG['ATTEMPTS']

    for _ in range(tentativas):
        try:
            element.click()
            break
        except ElementClickInterceptedException:
            time.sleep(2)
