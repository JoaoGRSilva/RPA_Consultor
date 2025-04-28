from openpyxl import load_workbook
import pandas as pd
import logging, os
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
                                   (df['CADASTRADO PLUXXE?'] == 'NÃO') & 
                                   (df['Processado RPA'] == 'NÃO')]
                                   
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
            return None

    @staticmethod
    def atualizar_esteira(resultados, caminho_excel):
        try:
            # Criar DataFrame com apenas as colunas necessárias
            df_novos = pd.DataFrame(resultados)
            df_novos.rename(columns={'numero': 'Contrato', 'status': 'Status'}, inplace=True)
            
            # Verificar se o arquivo existe
            if not os.path.exists(caminho_excel):
                df_novos.to_excel(
                    caminho_excel, 
                    engine='openpyxl', 
                    sheet_name='#RPA', 
                    index=False
                )
                return True
            
            # Se existe, atualizar ou adicionar novos registros
            try:
                df_existente = pd.read_excel(caminho_excel, sheet_name='#RPA')
                
                # Mesclar os dados existentes com os novos
                df_final = pd.concat([df_existente, df_novos]).drop_duplicates(subset=['Contrato'], keep='last')
                
                with pd.ExcelWriter(caminho_excel, engine='openpyxl', mode='w') as writer:
                    df_final[['Contrato', 'Status']].to_excel(
                        writer, 
                        sheet_name='#RPA', 
                        index=False
                    )
                    
                print(f"Planilha atualizada: {len(df_novos)} contratos")
                return True
                
            except Exception as e:
                df_novos[['Contrato', 'Status']].to_excel(
                    caminho_excel, 
                    engine='openpyxl', 
                    sheet_name='#RPA', 
                    index=False
                )
                return True

        except Exception as e:
            logging.error(f"Falha crítica na atualização: {str(e)}")
            if "File is not a zip file" in str(e):
                os.remove(caminho_excel)
                pd.DataFrame({'Contrato': [], 'Status': []}).to_excel(
                    caminho_excel, 
                    engine='openpyxl', 
                    sheet_name='#RPA',
                    index=False
                )
            return False

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