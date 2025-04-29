import time
from datetime import timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from config.config import CONFIG
from config.selectors import Selectors
from models.pdf_processor import PDFProcessor
from models.excel_processor import ExcelProcessor
from utils.helpers import *

class ContraktorBot:
    """Classe principal para automação do Contraktor."""

    def __init__(self, ui=None):
        self.driver = None
        self.ui = ui
        self.tempos_processamento = []
        self.tempo_inicio_total = None

    def iniciar_navegador(self):
        """Inicia o navegador Chrome."""
        try:
            options = Options()
            options.add_argument("--headless")  
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-gpu") 
            options.add_argument("--disable-dev-shm-usage") 
            self.driver = webdriver.Chrome(options=options)
            self.driver.maximize_window()
            self.driver.execute_script("document.title = 'Contraktor - Google Chrome'")
            return True
        except Exception as e:
            print(f"Erro ao iniciar navegador: {e}")
            return False

    def login(self):
        """Realiza login no sistema Contraktor."""
        try:
            self.driver.get(CONFIG['URL_LOGIN'])

            user_input = aguardar_elemento(self.driver, By.ID, Selectors.LOGIN_EMAIL)
            pwd_input = aguardar_elemento(self.driver, By.ID, Selectors.LOGIN_PASSWORD)

            user_input.send_keys(CONFIG['EMAIL'])
            pwd_input.send_keys(CONFIG['PASSWORD'])

            botao_login = aguardar_elemento(self.driver, By.XPATH, Selectors.LOGIN_BUTTON, tipo_espera='clicavel')
            try_click(botao_login)

            # Aguarda a mudança de URL para confirmar login
            WebDriverWait(self.driver, 10).until(EC.url_changes(CONFIG['URL_LOGIN']))
            print("Acesso ao Contracktor realizado com sucesso!")
            time.sleep(2)

            return True

        except Exception as e:
            print(f"Erro durante o login: {e}")
            return False

    def processar_contrato(self, numero_contrato, idx, total):
        """
        Processa um contrato específico.

        Args:
            numero_contrato: Número do contrato a ser processado
            idx: Índice atual do contrato
            total: Total de contratos

        Returns:
            dict: Resultado do processamento com status
        """
        arquivo_pdf = None

        try:
            print(f"\n=== Processando contrato {idx}/{total}: {numero_contrato} ===")

            # Atualizar UI
            if self.ui:
                self.ui.root.after(0, self.ui.atualizar_progresso, idx, total)

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
                print(f"Erro: Arquivo PDF não encontrado após download do contrato {numero_contrato}")
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato,
                    "status": "Falha",
                    "arquivo": None,
                    "erro": "Arquivo PDF não encontrado após download"
                }

            # Extrair informações do PDF
            pdf_info = PDFProcessor.extrair_informacoes(arquivo_pdf)
            dados_planilha = PDFProcessor.preparar_dados_para_planilha(pdf_info)

            if dados_planilha['Nome completo'] == "":
                excluir_arquivo(arquivo_pdf)
                print(f"Erro: Nome completo vazio no PDF do contrato {numero_contrato}")
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato,
                    "status": "Falha",
                    "arquivo": None,
                    "erro": "Nome completo vazio no PDF"
                }

            # Preencher planilha
            if dados_planilha is not None:
                ExcelProcessor.preencher_planilha(dados_planilha)
                excluir_arquivo(arquivo_pdf)
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato,
                    "status": "Sucesso",
                    "arquivo": arquivo_pdf
                }

        except Exception as e:
            excluir_arquivo(arquivo_pdf)
            try:
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)
            except:
                pass

            return {
                "numero": numero_contrato,
                "status": "Falha",
                "arquivo": None,
                "erro": str(e)
            }

    def executar(self, limite=None, modo_teste=False):
        """
        Executa o processamento completo de contratos.

        Args:
            limite: Número máximo de contratos a processar
            modo_teste: Se True, processa apenas o primeiro contrato
        """
        try:
            # Ler contratos pendentes
            contratos = ExcelProcessor.ler_contratos_pendentes(limite, modo_teste)

            if not contratos or contratos is None:
                print("Nenhum contrato pendente encontrado para processamento")
                return

            try:
                with open(CONFIG['EXCEL_PLUXXE'], 'r') as file:
                    print("Planilha pluxxe Ok!")

            except FileNotFoundError:
                print("Arquivo pluxxe não carregado, por favor carregar arquivo para pasta do robo!")
                return

            # Iniciar navegador
            if not self.iniciar_navegador():
                return

            time.sleep(2)

            # Realizar login
            if not self.login():
                return

            total_contratos = len(contratos)
            print(f"\n=== Processamento de {total_contratos} contratos iniciado ===")

            # Marca o tempo de início da execução total
            self.tempo_inicio_total = time.time()

            # Atualizar UI inicial
            if self.ui:
                self.ui.root.after(0, self.ui.atualizar_progresso, 0, total_contratos)

            # Processar cada contrato
            resultados = []
            for idx, numero_contrato in enumerate(contratos):
                # Marca tempo inicial do contrato
                tempo_inicio_contrato = time.time()

                # Mostrar estimativa se já tiver processado pelo menos um contrato
                if self.tempos_processamento and idx < total_contratos - 1:
                    tempo_medio = sum(self.tempos_processamento) / len(self.tempos_processamento)
                    tempo_restante = tempo_medio * (total_contratos - idx)

                    # Atualizar UI com tempo estimado
                    if self.ui:
                        self.ui.root.after(0, self.ui.atualizar_tempo_estimado, tempo_restante)

                resultado = self.processar_contrato(numero_contrato, idx + 1, total_contratos)
                resultados.append(resultado)  # Salva todos os resultados, incluindo falhas

                # Calcula o tempo que levou para processar este contrato
                tempo_contrato = time.time() - tempo_inicio_contrato
                self.tempos_processamento.append(tempo_contrato)

                # Pequena pausa entre contratos
                time.sleep(3)

            # Calcula o tempo total da execução
            tempo_total = time.time() - self.tempo_inicio_total
            tempo_total_formatado = str(timedelta(seconds=int(tempo_total)))

            # Resumo final
            sucesso = sum(1 for r in resultados if r['status'] == "Sucesso")

            # Atualiza a planilha com os resultados
            ExcelProcessor.atualizar_esteira(resultados, CONFIG['EXCEL_CONTRATOS'])

            contratos_erro = []
            for r in resultados:
                if r['status'] == "Falha":
                    contratos_erro.append(r['numero'])

            print(f"\n=== Resumo do processamento ===")
            print(f"Total de contratos processados: {len(resultados)}")
            print(f"Sucessos: {sucesso}")
            print(f"Falhas: {len(contratos_erro)}")
            print(f"Contratos com falha: {contratos_erro}")
            print(f"Tempo total de execução: {tempo_total_formatado}")

            if self.tempos_processamento:
                tempo_medio = sum(self.tempos_processamento)/len(self.tempos_processamento)
                print(f"Tempo médio por contrato: {str(timedelta(seconds=int(tempo_medio)))}")

        except Exception as e:
            print(f"Erro durante a execução: {e}")
        finally:
            if self.driver:
                self.driver.quit()
                print("Navegador fechado")

