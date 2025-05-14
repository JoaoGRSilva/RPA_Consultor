import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from config.config import CONFIG
from config.selectors import Selectors
from models.pdf_processor import PDFProcessor
from models.excel_processor import ExcelProcessor
from models.email_sender import EmailSender
from models.contracktor_processor import ContracktorProcessor
from utils.helpers import *
import warnings, shutil
from datetime import datetime,timedelta

from credentials import senha_contracktor, email_contracktor


warnings.filterwarnings("ignore")

class ContraktorBot:
    """Classe principal para automação do Contraktor."""

    def __init__(self, ui=None, excecao_var=None, log_queue=None):
        self.excecao_var = excecao_var
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

            user_input.send_keys(email_contracktor)
            pwd_input.send_keys(senha_contracktor)

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
                print(f"Erro: PDF fora do padrão: {numero_contrato}")
                botao_voltar = aguardar_elemento(self.driver, By.XPATH, Selectors.RETURN_BUTTON, tipo_espera='clicavel')
                try_click(botao_voltar)

                return {
                    "numero": numero_contrato,
                    "status": "Falha",
                    "arquivo": None,
                    "erro": "Erro: PDF fora do padrão:"
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

    def liberar_todas_as_fichas(self):
        processor = ContracktorProcessor(self.driver)
        while True:
            sucesso = processor.liberar_contratos()
            if not sucesso:
                print("✅Todas as fichas liberadas")
                break
            print("🔁Ficha liberada. Loopando para a próxima")
            time.sleep(1)



    def executar(self, limite=None, modo_teste=False):
        """
        Executa o processamento completo de contratos.

        Args:
            limite (int): Número máximo de contratos a processar.
            modo_teste (bool): Se True, processa apenas o primeiro contrato.
        """

        try:

            excecao = self.excecao_var.get()
            hoje = datetime.today().weekday()

            if hoje != 0 and not excecao:
                pass
            else:
                print("🔹 Iniciando processo de compilado...")
                compilado = ExcelProcessor.compilar_arquivos()
                print("✅ Compilação concluída.")
                ExcelProcessor.compilar_planilhas(compilado)
                print("📄 Planilha de compilamento gerada.")
                planilha_compilada = encontrar_excel_recente(CONFIG['COMPILADO_FOLDER'])
                EmailSender.envio_email(planilha_compilada)


            if modo_teste:
                print("⚠️ Modo de teste ativado. Encerrando após compilação.")
                return

            print("🔍 Lendo contratos pendentes...")
            contratos = ExcelProcessor.ler_contratos_pendentes(limite, modo_teste)

            if not contratos:
                print("⚠️ Nenhum contrato pendente encontrado.")
                return

            print("📁 Copiando modelo da planilha...")
            shutil.copy(CONFIG['EXCEL_PLUXXE'], 'planilha.xlsx')

            if not os.path.exists('planilha.xlsx'):
                print("❌ Planilha pluxxe não encontrada!")
                return
            else:
                print("✅ Planilha pluxxe carregada.")

            if not self.iniciar_navegador():
                return
            time.sleep(2)

            if not self.login():
                return

            try:
                self.liberar_todas_as_fichas()

            except Exception as e:
                print(f"Erro teste processar contrato: {e}")

            total_contratos = len(contratos)
            print(f"\n🚀 Iniciando processamento de {total_contratos} contratos...\n")
            self.tempo_inicio_total = time.time()

            if self.ui:
                self.ui.root.after(0, self.ui.atualizar_progresso, 0, total_contratos)

            resultados = []

            for idx, numero_contrato in enumerate(contratos):
                tempo_inicio_contrato = time.time()

                if self.tempos_processamento and idx < total_contratos - 1:
                    tempo_medio = sum(self.tempos_processamento) / len(self.tempos_processamento)
                    tempo_restante = tempo_medio * (total_contratos - idx)
                    if self.ui:
                        self.ui.root.after(0, self.ui.atualizar_tempo_estimado, tempo_restante)

                resultado = self.processar_contrato(numero_contrato, idx + 1, total_contratos)
                resultados.append(resultado)

                tempo_contrato = time.time() - tempo_inicio_contrato
                self.tempos_processamento.append(tempo_contrato)

                time.sleep(3)

            tempo_total = time.time() - self.tempo_inicio_total
            tempo_total_formatado = str(timedelta(seconds=int(tempo_total)))

            sucesso = sum(1 for r in resultados if r['status'] == "Sucesso")
            falhas = [r['numero'] for r in resultados if r['status'] == "Falha"]

            print("\n📊 === RESUMO DO PROCESSAMENTO ===")
            print(f"Total processado: {len(resultados)}")
            print(f"✅ Sucessos: {sucesso}")
            print(f"❌ Falhas: {len(falhas)}")
            print(f"Contratos com falha: {falhas}")
            print(f"⏱️ Tempo total: {tempo_total_formatado}")

            ExcelProcessor.renomear_saida()

            if self.tempos_processamento:
                tempo_medio = sum(self.tempos_processamento) / len(self.tempos_processamento)
                print(f"⏱️ Tempo médio por contrato: {str(timedelta(seconds=int(tempo_medio)))}")

        except Exception as e:
            print(f"❌ Erro durante a execução: {e}")

        finally:
            if self.driver:
                self.driver.quit()
                print("🧹 Navegador fechado.")
