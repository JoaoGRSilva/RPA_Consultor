import PyPDF2
import re
from utils import *

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
        info = {}

        try:
            with open(caminho_pdf, 'rb') as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)
                texto_completo = "".join([pagina.extract_text() + "\n" for pagina in leitor.pages])

                campos_busca = [
                    "CPF/MF", "Data de Nascimento", "Nome completo", "Celular", 
                    "(e-mail pessoal ou pessoal corporativo)", "RG", "Órgão Emissor", 
                    "Logradouro", "Número", "Bairro", "Cidade", "Estado", "CEP", "Celular"
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
                                info["Complemento"] = complemento
                            else:
                                info["Número"] = valor 
                        else:
                            info[campo] = valor
                            
                return {
                    "texto_completo": texto_completo, 
                    "dados_extraidos": info, 
                    "status": "Sucesso"
                }

        except Exception as e:
            print(f"Erro ao processar o PDF: {e}")
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
        # Mapeamento de abreviações
        abreviacoes_logradouro = {
            "Rua": "R.",
            "Avenida": "Av.",
            "Travessa": "Tv.",
            "Alameda": "Al.",
            "Praça": "Pç.",
            "Rodovia": "Rod.",
            "Estrada": "Est.",
            "Viela": "Vl.",
            "Ladeira": "Ld.",
            "Largo": "Lg.",
            "Beco": "Bco.",
            "Vila": "Vl.",
            "Conjunto": "Cj.",
            "Quadra": "Qd.",
            "Setor": "St.",
            "Loteamento": "Lt.",
            "Caminho": "Cam.",
            "Servidão": "Serv."
        }

        dados_planilha = {
            'cod pluxe': "3230687",
            'sit ben': "Ativo",
            'Nome completo': "",
            'CPF/MF': "",
            'Data de Nascimento': "",
            'nome gravacao': "",
            'pular1': None,
            'pular2': None,
            'pular3': None,
            'pular4': None,
            'tipo': "023 - Pedido de 1ªVia de Cartão Sem Crédito",
            'produto': "603903 - Carteira Gift",
            'pular5': None,
            'pular6': None,
            'pular7': None,
            'local': "MATRIZ",
            'CEP': "",
            'Logradouro': "",
            'Número': "",
            'Complemento': "",
            'pular9': None,
            'Bairro': "",
            'Cidade': "",
            'Estado': "",
            'Responsável pelo recebimento': "",
            'DDD': None,
            'Celular': None,
            '(e-mail pessoal ou pessoal corporativo)': ""
        }

        if dados_pdf and dados_pdf['dados_extraidos']:
            dados_planilha.update({
                k: v for k, v in dados_pdf['dados_extraidos'].items() if k in dados_planilha
            })
            dados_planilha['Estado'] = estado_para_uf(dados_planilha['Estado'])

            if 'Nome completo' in dados_planilha and dados_planilha['Nome completo']:
                nome = dados_planilha['Nome completo']
                nome_abreviado = abreviar_nome(nome)
                dados_planilha["nome gravacao"] = nome_abreviado
                dados_planilha["Responsável pelo recebimento"] = nome_abreviado

            if 'Celular' in dados_planilha and dados_planilha['Celular']:
                telefone = dados_planilha['Celular'].strip()
                ddd, numero = split_numero(telefone)
                dados_planilha['DDD'] = ddd
                dados_planilha['Celular'] = numero

            if 'CEP' in dados_planilha and dados_planilha['CEP']:
                dados_planilha['CEP'] = dados_planilha['CEP'].replace("-", "")

            if 'Complemento' in dados_planilha and not dados_planilha['Complemento']:
                dados_planilha['Complemento'] = None

            logradouro = dados_planilha.get('Logradouro', '')
            logradouro_strip = logradouro.strip()

            for tipo_extenso, abrev in abreviacoes_logradouro.items():
                if logradouro_strip.lower().startswith(tipo_extenso.lower() + " "):
                    restante = logradouro_strip[len(tipo_extenso):].strip()
                    restante_formatado = restante.title()
                    dados_planilha['Logradouro'] = f"{abrev} {restante_formatado}"
                    break


            # Limpa espaços indevidos nos demais campos
            for chave, valor in dados_planilha.items():
                if chave not in ['Nome completo', 'nome gravacao', 'Responsável pelo recebimento', 'tipo', 'produto', 'Bairro', 'Cidade', 'Logradouro'] and isinstance(valor, str):
                    dados_planilha[chave] = ''.join(valor.split())

        return dados_planilha
