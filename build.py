# build.py
import os
import sys
import shutil
import subprocess
import PyInstaller.__main__

def build_executable():
    """Gera o executável usando PyInstaller"""
    print("Construindo executável...")

    # Caminho do .env formatado corretamente para --add-data
    env_data_arg = '.env;.'

    # Opções do PyInstaller
    options = [
        'main.py',                      # Script principal
        '--name=ContraktorBot',         # Nome do executável
        '--onefile',                    # Arquivo único
        '--windowed',                   # Sem janela de console
        f'--add-data={env_data_arg}',   # Incluir .env no executável
        '--hidden-import=selenium',     # Importações ocultas
        '--hidden-import=openpyxl',
        '--clean',                      # Limpar cache
    ]

    # Executar PyInstaller
    PyInstaller.__main__.run(options)

    print("Copiando arquivos adicionais...")

    # Copiar arquivos adicionais para o diretório dist
    additional_files = [
        'README.md',  # Adicione outros arquivos se necessário
    ]

    dist_dir = os.path.join('dist')
    exe_dir = os.path.join(dist_dir, 'ContraktorBot')

    os.makedirs(exe_dir, exist_ok=True)

    for file in additional_files:
        if os.path.exists(file):
            shutil.copy2(file, os.path.join(exe_dir))

    print("Build concluído com sucesso!")
    print("O executável está disponível em: dist/ContraktorBot.exe")

def install_requirements():
    """Instala as dependências necessárias"""
    print("Instalando dependências...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])

if __name__ == "__main__":
    # Instalar dependências
    install_requirements()

    # Construir o executável
    build_executable()
