from utils.helpers import *
from config.selectors import Selectors

from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
from time import sleep

class ContracktorProcessor:

    def __init__(self, driver):
        self.driver = driver

    def ajustar_contrato(self):
        try:
            ficha = aguardar_elemento(self.driver, By.XPATH, '//*[@id="EditorWrapper"]/div/div/div/div/div[2]/table[3]/tbody/tr[1]/td[2]/div', tipo_espera='presenca')
            cnpj = ficha.text
            titulo_ficha = f"{cnpj}, 99999,VENDEDOR"

            titulo = aguardar_elemento(self.driver, By.XPATH, Selectors.TITLE_INPUT)
            titulo.clear()
            titulo.send_keys(titulo_ficha)

        except:
            pass 

        try:
            botao_status = aguardar_elemento(self.driver, By.XPATH, Selectors.STATUS_BUTTON, tipo_espera='clicavel')
            try_click(botao_status)

            botao_vigente = aguardar_elemento(self.driver, By.XPATH, Selectors.VIGENTE_BUTTON, tipo_espera='clicavel')
            try_click(botao_vigente)

            botao_salvar = aguardar_elemento(self.driver, By.XPATH, Selectors.SALVAR_BUTTON, tipo_espera='clicavel')
            try_click(botao_salvar)

            contratos_botao = aguardar_elemento(self.driver, By.XPATH, Selectors.CONTRATOS_BUTTON, tipo_espera='clicavel')
            try_click(contratos_botao)

            formularios = aguardar_elemento(self.driver, By.XPATH, Selectors.FORMULARIO, tipo_espera='clicavel')
            try_click(formularios)

        except:
            pass 

    def liberar_contratos(self):
        liberou = False

        try:
            print("🔄 Acessando a aba de formulários...")
            formularios = aguardar_elemento(self.driver, By.XPATH, Selectors.FORMULARIO, tipo_espera='clicavel')
            try_click(formularios)

            print("⌛ Aguardando tabela de contratos aparecer...")
            sleep(2)
            aguardar_elemento(self.driver, By.XPATH, Selectors.TABLE, tipo_espera='presenca')

            rows = self.driver.find_elements(By.XPATH, "//tr[@role='row']")
            #print(f"📄 Total de linhas encontradas: {len(rows)}")

            for idx, row in enumerate(rows):
                if not row.is_displayed():
                    #print(f"🔕 Linha {idx} oculta, ignorando.")
                    continue

                try:
                    cells = row.find_elements(By.XPATH, ".//td[@role='cell']")
                    if not cells:
                        #print(f"⚠️ Linha {idx} sem células, ignorando.")
                        continue

                    texts = [cell.text.strip().lower() for cell in cells]
                    #print(f"🧪 Linha {idx} - Conteúdo: {texts}")

                    if any("ficha consultor" in text for text in texts):
                        #print(f"✅ Ficha consultor encontrada na linha {idx}. Liberando...")

                        botao_mais = row.find_element(By.XPATH, Selectors.OPTIONS_BUTTON)
                        try_click(botao_mais)

                        botao_gerar = aguardar_elemento(self.driver, By.XPATH, Selectors.GERAR_BUTTON, tipo_espera='clicavel')
                        try_click(botao_gerar)

                        self.ajustar_contrato()

                        liberou = True
                        return True

                except StaleElementReferenceException:
                    print(f"♻️ Elemento obsoleto na linha {idx}, tentando a próxima.")
                    continue

                except Exception as e:
                    print(f"❌ Erro inesperado na linha {idx}: {e}")
                    continue

        except Exception as e:
            print(f"❌ Erro ao acessar ou processar a tabela de formulários: {e}")

        if not liberou:
            print("📭 Nenhuma ficha disponível para liberar.")
        return liberou