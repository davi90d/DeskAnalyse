"""
Script para criar um executável da aplicação usando PyInstaller.
Configurado para resolver problemas de múltiplas janelas e falha de carregamento.
"""

import os
import sys
import subprocess
import platform

def create_spec_file():
    """Cria um arquivo .spec personalizado para o PyInstaller."""
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DiagnosticoHardware',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
    uac_admin=True,  # Solicita privilégios de administrador
)
"""
    
    with open("diagnostico_hardware.spec", "w") as f:
        f.write(spec_content)
    
    print("Arquivo .spec criado com sucesso.")

def run_pyinstaller():
    """Executa o PyInstaller com as configurações adequadas."""
    # Verifica se o PyInstaller está instalado
    try:
        subprocess.run(["pyinstaller", "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except (subprocess.SubprocessError, FileNotFoundError):
        print("PyInstaller não encontrado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # Cria o arquivo .spec personalizado
    create_spec_file()
    
    # Executa o PyInstaller usando o arquivo .spec
    print("Executando PyInstaller...")
    subprocess.run(["pyinstaller", "--clean", "diagnostico_hardware.spec"], check=True)
    
    print("\nExecutável criado com sucesso!")
    print(f"Localização: {os.path.abspath('dist/DiagnosticoHardware.exe')}")

if __name__ == "__main__":
    if platform.system() != "Windows":
        print("Este script deve ser executado em um ambiente Windows.")
        sys.exit(1)
    
    run_pyinstaller()
