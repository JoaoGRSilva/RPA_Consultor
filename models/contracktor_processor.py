from utils.helpers import *
from config.selectors import Selectors

from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
import time

class ContracktorProcessor:

    def __init__(self, driver):
        self.driver = driver

    def ajustar_contrato(self):
        try:
            ficha = aguardar_elemento(self.driver, By.XPATH, '//*[@id="EditorWrapper"]/div/div/div/div/div[2]/table[3]/tbody/tr[1]/td[2]/div', tipo_espera='presenca')
            cnpj = ficha.text
            titulo_ficha = f"{cnpj}, VENDEDOR"

            titulo = aguardar_elemento(self.driver, By.XPATH, Selectors.TITLE_INPUT)
            titulo.clear()
            titulo.send_keys(titulo_ficha)

        except:
            pass  # Erro ao renomear ficha

        try:
            botao_status = aguardar_elemento(self.driver, By.XPATH, Selectors.STATUS_BUTTON, tipo_espera='clicavel')
            try_click(botao_status)

            botao_vigente = aguardar_elemento(self.driver, By.XPATH, Selectors.VIGENTE_BUTTON, tipo_espera='clicavel')
            try_click(botao_vigente)

            contratos_botao = aguardar_elemento(self.driver, By.XPATH, Selectors.CONTRATOS_BUTTON, tipo_espera='clicavel')
            try_click(contratos_botao)

            formularios = aguardar_elemento(self.driver, By.XPATH, Selectors.FORMULARIO, tipo_espera='clicavel')
            try_click(formularios)

        except:
            pass  # Erro ao alterar status

    def liberar_contratos(self):
        try:
            formularios = aguardar_elemento(self.driver, By.XPATH, Selectors.FORMULARIO, tipo_espera='clicavel')
            try_click(formularios)

            aguardar_elemento(self.driver, By.XPATH, Selectors.TABLE, tipo_espera='presenca')

            try:
                container = self.driver.find_element(By.XPATH, "//div[contains(@class, 'MuiTableContainer')]")
                self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", container)
                time.sleep(1)
            except:
                pass 

            rows = self.driver.find_elements(By.XPATH, "//tr[@role='row']")

            for idx, row in enumerate(rows):
                if not row.is_displayed():
                    continue

                try:
                    cells = row.find_elements(By.XPATH, ".//td[@role='cell']")
                    if not cells:
                        continue

                    texts = [cell.text.strip().lower() for cell in cells]

                    if any("ficha consultor" in text for text in texts):
                        botao_mais = row.find_element(By.XPATH, ".//button[contains(@aria-label, 'Opções')]")
                        try_click(botao_mais)

                        botao_gerar = aguardar_elemento(self.driver, By.XPATH, Selectors.GERAR_BUTTON, tipo_espera='clicavel')
                        try_click(botao_gerar)

                        self.ajustar_contrato()
                        return True

                except StaleElementReferenceException:
                    continue  

                except:
                    continue

        except:
            pass 
