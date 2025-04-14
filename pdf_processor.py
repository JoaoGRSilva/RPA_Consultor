import PyPDF2
import logging
import re
from utils import abreviar_nome, estado_para_uf

class PDFProcessor:
    """Classe para processar arquivos PDF de contratos."""
    
    @staticmethod
    def extrair_informacoes(caminho_pdf):
        """
        Extrai informações relevantes de um PDF de contrato.
        
        Args:
            caminho_pdf: Caminho para o arquivo PDF
            
        Returns:
            dict: Dicionário com informações extraídas e status
        """
        print(f"Extraindo informações do PDF: {caminho_pdf}")
        info = {}

        try:
            with open(caminho_pdf, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                texto_completo = "".join([pagina.extract_text() + "\n" for pagina in leitor.pages])

                campos_busca = [
                    "CPF/MF", "Data de Nascimento", "Nome completo", "Celular", 
                    "(e-mail pessoal ou pessoal corporativo)", "RG", "Órgão Emissor", 
                    "Logradouro", "Número", "Bairro", "Cidade", "Estado", "CEP"
                ]

                for campo in campos_busca:
                    if campo in texto_completo:
                        indice = texto_completo.find(campo)
                        valor = texto_completo[indice + len(campo):].split('\n')[0].strip()

                        if campo == "Número":
                            match = re.match(r'(\d+)\s*(.*)', valor)
                            if match:
                                numero = match.group(1)
                                complemento = match.group(2).strip()
                                info["Número"] = numero
                                info["pular8"] = complemento
                                print(f"Encontrado: Número - {numero}, Complemento - {complemento}")
                            else:
                                info["Número"] = valor
                                print(f"Encontrado: Número - {valor}")
                        else:
                            info[campo] = valor
                            print(f"Encontrado: {campo} - {valor}")

                return {
                    "texto_completo": texto_completo, 
                    "dados_extraidos": info, 
                    "status": "Sucesso"
                }

        except Exception as e:
            logging.error(f"Erro ao processar o PDF: {e}")
            return {
                "texto_completo": "", 
                "dados_extraidos": {}, 
                "status": f"Erro: {str(e)}"
            }
    
    @staticmethod
    def preparar_dados_para_planilha(dados_pdf):
        """
        Prepara os dados extraídos do PDF para inserção na planilha.
        
        Args:
            dados_pdf: Dicionário com dados extraídos do PDF
                
        Returns:
            dict: Dados formatados para planilha na ordem correta
        """
        # Inicializa o dicionário com campos vazios na ordem correta
        dados_planilha = {
            'cod pluxe': "3230687",                    # Código Cliente Pluxee
            'sit ben': "Ativo",                        # Situação do beneficiário
            'Nome completo': "",                       # Nome completo
            'CPF/MF': "",                                 # CPF
            'Data de Nascimento': "",                  # Data de nascimento
            'nome gravacao': "",                       # Nome para gravação no cartão
            'pular1': None,                            # Pula 1
            'pular2': None,                            # Pula 2
            'pular3': None,                            # Pula 3
            'pular4': None,                            # Pula 4
            'tipo': "023 - Pedido de 1ªVia de Cartão Sem Crédito",  # Tipo do pedido
            'produto': "603903 - Carteira Gift",       # Produto
            'pular5': None,                            # Pula 5
            'pular6': None,                            # Pula 6
            'pular7': None,                            # Pula 7
            'local': "MATRIZ",                         # Local de entrega
            'CEP': "",                                 # CEP
            'Logradouro': "",                            # Endereço
            'Número': "",                              # Número
            'Complemento': "",                         # Complemento (pula 8 quando não tem)
            'pular9': None,                            # Pula 9
            'Bairro': "",                              # Bairro
            'Cidade': "",                              # Cidade
            'Estado': "",                                  # UF
            'Responsável pelo recebimento': "",        # Responsável pelo recebimento
            'pular10': None,                           # Pula 10
            'pular11': None,                           # Pula 11
            '(e-mail pessoal ou pessoal corporativo)': ""  # Email do beneficiário
        }
        
        # Atualiza com os dados extraídos do PDF
        if dados_pdf and dados_pdf['dados_extraidos']:
            dados_planilha.update({
                k: v for k, v in dados_pdf['dados_extraidos'].items() if k in dados_planilha
            })
            dados_planilha['Estado'] = estado_para_uf(dados_planilha['Estado'])
            
            # Processar nome para abreviação
            if 'Nome completo' in dados_planilha and dados_planilha['Nome completo']:
                nome = dados_planilha['Nome completo']
                nome_abreviado = abreviar_nome(nome)
                dados_planilha["nome gravacao"] = nome_abreviado
                dados_planilha["Responsável pelo recebimento"] = nome_abreviado
                
            # Verifica se Complemento está vazio e o substitui por None (pula 8)
            if 'Complemento' in dados_planilha and not dados_planilha['Complemento']:
                dados_planilha['Complemento'] = None
        
        return dados_planilha