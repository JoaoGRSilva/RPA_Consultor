from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import logging
import time
from selenium.webdriver.support.ui import WebDriverWait
from utils import aguardar_elemento, encontrar_arquivo_recente, excluir_arquivo, try_click
from config import CONFIG
from selectors_1 import Selectors
from pdf_processor import PDFProcessor
from excel_processor import ExcelProcessor

class ContraktorBot:
    """Classe principal para automação do Contraktor."""
    
    def __init__(self):
        self.driver = None
        self.setup_logging()
        
    def setup_logging(self):
        """Configura o sistema de logging."""
        logging.basicConfig(
            filename=CONFIG['LOG_FILE'],
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        # Adicionar também log no console
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        console.setFormatter(logging.Formatter('%(message)s'))
        logging.getLogger('').addHandler(console)
        
    def iniciar_navegador(self):
        """Inicia o navegador Chrome."""
        try:
            self.driver = webdriver.Chrome()
            self.driver.maximize_window()
            print("Navegador Chrome iniciado com sucesso")
            return True
        except Exception as e:
            print(f"Erro ao iniciar navegador: {e}")
            return False
            
    def login(self):
        """Realiza login no sistema Contraktor."""
        try:
            print("Realizando login no Contraktor...")
            self.driver.get(CONFIG['URL_LOGIN'])

            user_input = aguardar_elemento(self.driver, By.ID, Selectors.LOGIN_EMAIL)
            pwd_input = aguardar_elemento(self.driver, By.ID, Selectors.LOGIN_PASSWORD)

            user_input.send_keys(CONFIG['EMAIL'])
            pwd_input.send_keys(CONFIG['PASSWORD'])

            botao_login = aguardar_elemento(self.driver, By.XPATH, Selectors.LOGIN_BUTTON, tipo_espera='clicavel')
            try_click(botao_login)

            # Aguarda a mudança de URL para confirmar login
            WebDriverWait(self.driver, 10).until(EC.url_changes(CONFIG['URL_LOGIN']))
            print("Login realizado com sucesso!")
            print("Sleep para o primeiro contrato...")
            time.sleep(2)
            return True
            
        except Exception as e:
            logging.error(f"Erro durante o login: {e}")
            return False
            
    def processar_contrato(self, numero_contrato):
        """
        Processa um contrato específico.
        
        Args:
            numero_contrato: Número do contrato a ser processado
            
        Returns:
            dict: Resultado do processamento com status
        """
        try:
            print(f"\n=== Processando contrato: {numero_contrato} ===")

            # Aguardar carregamento da página
            aguardar_elemento(self.driver, By.XPATH, Selectors.LOADING_SPINNER, tipo_espera='invisibilidade')

            # Selecionar tipo de busca (Número)
            dropdown = aguardar_elemento(self.driver, By.XPATH, Selectors.SEARCH_DROPDOWN, tipo_espera='clicavel')
            try_click(dropdown)

            opcao_numero = aguardar_elemento(self.driver, By.XPATH, Selectors.SEARCH_OPTION_NUMERO, tipo_espera='clicavel')
            try_click(opcao_numero)

            # Preencher número de contrato
            campo_busca = aguardar_elemento(self.driver, By.XPATH, Selectors.SEARCH_INPUT)
            campo_busca.clear()
            campo_busca.send_keys(numero_contrato)

            # Clicar no botão de busca
            botao_busca = aguardar_elemento(self.driver, By.XPATH, Selectors.SEARCH_BUTTON, tipo_espera='clicavel')
            try_click(botao_busca)

            # Clicar no link do contrato
            link_contrato = aguardar_elemento(self.driver, By.XPATH, Selectors.CONTRACT_LINK, tipo_espera='clicavel')
            try_click(link_contrato)

            # Clicar no botão de download
            botao_download = aguardar_elemento(self.driver, By.XPATH, Selectors.DOWNLOAD_BUTTON, tipo_espera='clicavel')
            try_click(botao_download)

            # Encontrar arquivo baixado
            arquivo_pdf = encontrar_arquivo_recente(CONFIG['DOWNLOAD_FOLDER'])

            if arquivo_pdf is None:
                logging.error("Arquivo PDF não encontrado após o download")
                
                # Voltar para a tela de busca
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato, 
                    "status": "Falha", 
                    "arquivo": None
                }

            # Extrair informações do PDF
            pdf_info = PDFProcessor.extrair_informacoes(arquivo_pdf)
            
            # Preparar dados para a planilha
            dados_planilha = PDFProcessor.preparar_dados_para_planilha(pdf_info)

            if dados_planilha['Nome completo'] == "":
                # Excluir arquivo PDF
                excluir_arquivo(arquivo_pdf)
                
                # Voltar para a tela de busca
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                "numero": numero_contrato, 
                "status": "Falha", 
                "arquivo": None
            }


            if dados_planilha is not None:     
                # Preencher planilha
                ExcelProcessor.preencher_planilha(dados_planilha)
                
                # Excluir arquivo PDF
                excluir_arquivo(arquivo_pdf)
                
                # Voltar para a tela de busca
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato, 
                    "status": "Sucesso", 
                    "arquivo": arquivo_pdf
                }

        except Exception as e:
            logging.error(f"Erro ao processar contrato {numero_contrato}: {e}")

            # Excluir arquivo PDF
            excluir_arquivo(arquivo_pdf)
            
            # Voltar para a tela de busca
            botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
            try_click(botao_voltar)            

            return {
                "numero": numero_contrato, 
                "status": "Falha", 
                "arquivo": None
            }
            
    def executar(self, limite=None, modo_teste=False):
        """
        Executa o processamento completo de contratos.
        
        Args:
            limite: Número máximo de contratos a processar
            modo_teste: Se True, processa apenas o primeiro contrato
        """
        try:
            # Iniciar navegador
            if not self.iniciar_navegador():
                return
                
            # Realizar login
            if not self.login():
                return
                
            # Ler contratos pendentes
            contratos = ExcelProcessor.ler_contratos_pendentes(limite, modo_teste)
            
            if not contratos:
                logging.warning("Nenhum contrato pendente encontrado para processamento")
                return

            # Processar cada contrato
            resultados = []
            for idx, numero_contrato in enumerate(contratos):
                logging.info(f"Processando contrato {idx + 1}/{len(contratos)}: {numero_contrato}")
                resultado = self.processar_contrato(numero_contrato)
                resultados.append(resultado)
                
                # Pequena pausa entre contratos
                time.sleep(3)
                
            # Resumo final
            sucesso = sum(1 for r in resultados if r['status'] == "Sucesso")

            contratos_erro = []
            for r in resultados:
                if r['status'] == "Falha":
                    contratos_erro.append(r['numero'])

            print(f"\n=== Resumo do processamento ===")
            print(f"Total de contratos processados: {len(resultados)}")
            print(f"Sucessos: {sucesso}")
            print(f"Falhas: {len(resultados) - sucesso}")
            print(f"Contratos com falha: {contratos_erro}")
            
        except Exception as e:
            logging.error(f"Erro durante a execução: {e}")
        finally:
            if self.driver:
                self.driver.quit()
                print("Navegador fechado")