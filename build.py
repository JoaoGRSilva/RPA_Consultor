# build.py
import os
import sys
import shutil
import subprocess
import PyInstaller.__main__

def clean_build_directories():
    """Limpa os diretórios de build anteriores"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    for file in spec_files:
        os.remove(file)

def build_executable():
    """Gera o executável usando PyInstaller"""
    print("Construindo executável...")
    
    # Opções do PyInstaller
    options = [
        'main.py',                      # Script principal
        '--name=ContraktorBot',         # Nome do executável
        '--onefile',                    # Arquivo único
        '--windowed',                   # Sem janela de console
        '--hidden-import=selenium',     # Importações ocultas
        '--hidden-import=openpyxl',
        '--clean',                      # Limpar cache
    ]
    
    # Executar PyInstaller
    PyInstaller.__main__.run(options)
    
    print("Copiando arquivos adicionais...")
    # Copiar arquivos adicionais para o diretório dist
    additional_files = [
        '.env',                         # Configurações de ambiente
        'README.md',                    # Instruções
        # Adicione outros arquivos necessários aqui
    ]
    
    for file in additional_files:
        if os.path.exists(file):
            shutil.copy2(file, os.path.join('dist', 'ContraktorBot'))
    
    print("Build concluído com sucesso!")
    print("O executável está disponível em: dist/ContraktorBot/ContraktorBot.exe")

def install_requirements():
    """Instala as dependências necessárias"""
    print("Instalando dependências...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])

if __name__ == "__main__":
    # Instalar dependências
    install_requirements()
    
    # Limpar diretórios de build anteriores
    clean_build_directories()
    
    # Construir o executável
    build_executable()