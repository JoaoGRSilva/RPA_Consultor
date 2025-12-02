# 🤖 Módulo de Execução Automatizada

## 📌 Sobre o projeto

Este projeto é um sistema de **RPA (Robotic Process Automation)** desenvolvido em **Python** que automatiza a extração de informações da plataforma **Contracktor** e as organiza em uma planilha estruturada, pronta para integração com outros sistemas.

---

## 🎯 Objetivo

Automatizar o fluxo de trabalho de coleta de dados da Contracktor, **eliminando extrações manuais**, **reduzindo erros humanos** e **aumentando a eficiência operacional** por meio de um processo padronizado de transferência de dados entre sistemas.

---

## 💡 Funcionalidades

- 🔍 Extração automatizada de dados da plataforma Contracktor  
- 🧠 Processamento e formatação inteligente das informações  
- 📊 Preenchimento automático de planilhas em formato específico  
- 🔗 Integração facilitada com sistemas de destino  
- 🖱️ Interface simples para execução da automação

---

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **Selenium** – Automação de navegação web  
- **Pandas** – Manipulação de dados  
- **PyMuPDF & PyPDF2** – Leitura e processamento de arquivos PDF  
- **Openpyxl** – Manipulação de planilhas Excel

---

## 📦 Requisitos

As dependências estão listadas no arquivo `requirements.txt`. Principais bibliotecas:

selenium==4.29.0 pandas==2.2.3 openpyxl==3.1.5 PyMuPDF==1.25.4 PyPDF2==3.0.1 python-dotenv==1.1.0

yaml
Copiar
Editar

---

## 💻 Instalação

Você pode rodar este projeto de duas formas:

### 1. Executando o código-fonte

```bash
# Clonar o repositório
git clone https://github.com/JoaoGRSilva/RPA_Consultor
cd modulo-execucao-automatizada

# Instalar as dependências
pip install -r requirements.txt

# Executar o script principal
python main.py
```

2. Utilizando o executável
Faça o download da última versão na aba Releases

Execute o arquivo diretamente (sem necessidade de instalar dependências)

🗂️ Estrutura do Projeto
```plaintext
.
├── config/
│   ├── config.py
│   └── selectors.py
├── controllers/
│   └── contraktor_bot.py
├── dist/
├── env/
├── models/
│   ├── excel_processor.py
│   └── pdf_processor.py
├── utils/
│   └── helpers.py
├── views/
│   └── contraktor_ui.py
├── .env
├── .gitignore
├── build.py
├── erros_contratos.txt
├── main.py
├── piloto.xlsx
└── requirements.txt
```

⚙️ Como Usar
Configure as credenciais e parâmetros no arquivo .env e nos arquivos de configuração.

Execute main.py ou o executável.

O sistema irá:

🌐 Acessar a plataforma Contracktor

📥 Extrair os dados necessários

🧹 Processar e formatar os dados

📤 Gerar uma planilha compatível com o sistema de destino

🧪 Notas de Desenvolvimento
Este projeto respeita as políticas de uso da plataforma Contracktor, utilizando técnicas de web scraping ético. Mantenha as dependências sempre atualizadas para garantir a compatibilidade com mudanças na plataforma.

📄 Licença
Distribuído sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.
