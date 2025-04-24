from openpyxl import load_workbook
import pandas as pd
import logging
from config import CONFIG

class ExcelProcessor:
    """Classe para processar planilhas relacionadas aos contratos."""

    @staticmethod
    def ler_contratos_pendentes(limite=None, modo_teste=False):
        try:
            caminho_excel = CONFIG['EXCEL_ESTEIRA']
            print(f"Lendo planilha de contratos: {caminho_excel}")
            
            df = pd.read_excel(caminho_excel, sheet_name='CONSOLIDADO', header=1)
            
            contratos_vigentes = df[(df['STATUS'] == 'Vigente') & 
                                   (df['Status Omni'] == 'Demandado Para Logística') & 
                                   (df['CADASTRADO PLUXXE?'] == 'NÃO')]
                                   
            contratos_lista = contratos_vigentes['Numeração'].tolist()
            
            total = len(contratos_lista)
            print(f"Total de {total} contratos pendentes encontrados")
            
            if modo_teste:
                contratos_lista = contratos_lista[:1]
                logging.info("Modo de teste ativado - processando apenas 1 contrato")
            elif limite and limite > 0:
                contratos_lista = contratos_lista[:limite]
                logging.info(f"Limite aplicado - processando {len(contratos_lista)} contratos")
                
            return contratos_lista
            
        except Exception as e:
            logging.error(f"Erro ao ler planilha de contratos: {e}")
            return []

    # Na classe ExcelProcessor
    @staticmethod
    def atualizar_esteira(resultados, caminho_excel):
        """
        Atualiza a planilha de esteira com os resultados do processamento.
        
        Args:
            resultados: Lista de dicionários com os resultados do processamento
            caminho_excel: Caminho para a planilha de esteira
        """
        try:
            # Extrair apenas os dados necessários para a planilha
            dados_planilha = []
            for r in resultados:
                dados_planilha.append({
                    'Numero': r['numero'],
                    'Status': r['status']
                })
                
            df = pd.DataFrame(dados_planilha)
            
            with pd.ExcelWriter(caminho_excel, engine='openpyxl', mode='a', 
                            if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='#RPA', index=False)
                
            print(f"Planilha de esteira atualizada com {len(dados_planilha)} contratos")
            
        except Exception as e:
            logging.error(f"Erro ao atualizar a esteira: {e}")

    @staticmethod
    def preencher_planilha(dados, linha_inicial=7):
        try:
            pluxxe_path = CONFIG['EXCEL_PLUXXE']
            wb = load_workbook(pluxxe_path)
            sheet = wb['Dados dos Beneficiários']

            linha_atual = linha_inicial
            while sheet.cell(row=linha_atual, column=1).value is not None:
                linha_atual += 1

            for coluna, (chave, valor) in enumerate(dados.items(), start=1):
                sheet.cell(row=linha_atual, column=coluna, value=valor)

            wb.save(pluxxe_path)
            print(f"Dados inseridos na linha {linha_atual} da planilha PLUXXE")            
            return True

        except Exception as e:
            logging.error(f"Erro ao preencher a planilha: {e}")
            return False