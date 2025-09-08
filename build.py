import os
import sys
import shutil
import subprocess
import PyInstaller.__main__

APP_NAME = "ContraktorBot"

def build_executable():
    """Gera o executável usando PyInstaller"""
    print("🚀 Construindo executável...")

    options = [
        'main.py',
        f'--name={APP_NAME}',
        '--onefile',
        '--windowed',
        '--hidden-import=selenium',
        '--hidden-import=openpyxl',
        '--clean',
        # '--icon=icon.ico',  # opcional
    ]

    PyInstaller.__main__.run(options)

    print("📂 Copiando arquivos adicionais...")

    additional_files = ['README.md']
    dist_dir = os.path.join('dist')

    for file in additional_files:
        if os.path.exists(file):
            shutil.copy2(file, dist_dir)

    print("✅ Build concluído com sucesso!")
    print(f"📦 Executável disponível em: dist/{APP_NAME}.exe")

def install_requirements():
    """Instala as dependências necessárias"""
    print("📦 Instalando dependências...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])

if __name__ == "__main__":
    if "--install" in sys.argv:
        install_requirements()
    build_executable()
