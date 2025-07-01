#!/usr/bin/env python3
"""
Script para construir o executável otimizado
"""
import os
import sys
import subprocess
import shutil
from pathlib import Path

def main():
    print("=== CONSTRUTOR DE EXECUTÁVEL OTIMIZADO ===")
    print("Versão 2.0 - Automatização de Reports")
    print("-" * 50)
    
    # Verificar se o PyInstaller está instalado
    try:
        import PyInstaller
        print(f"✓ PyInstaller encontrado: {PyInstaller.__version__}")
    except ImportError:
        print("✗ PyInstaller não encontrado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("✓ PyInstaller instalado com sucesso!")
    
    # Verificar se o openpyxl está instalado
    try:
        import openpyxl
        print(f"✓ openpyxl encontrado: {openpyxl.__version__}")
    except ImportError:
        print("✗ openpyxl não encontrado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], check=True)
        print("✓ openpyxl instalado com sucesso!")
    
    # Limpar builds anteriores
    print("\nLimpando builds anteriores...")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"✓ Pasta {folder} removida")
    
    # Verificar se o arquivo principal existe
    if not os.path.exists("main_optimized.py"):
        print("✗ Arquivo main_optimized.py não encontrado!")
        return False
    
    # Verificar se o arquivo .spec existe
    if not os.path.exists("main_optimized.spec"):
        print("✗ Arquivo main_optimized.spec não encontrado!")
        return False
    
    print("\nIniciando compilação...")
    print("Isso pode levar alguns minutos...")
    
    try:
        # Executar PyInstaller
        result = subprocess.run([
            sys.executable, "-m", "PyInstaller", 
            "main_optimized.spec",
            "--clean",
            "--noconfirm"
        ], capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✓ Compilação concluída com sucesso!")
            
            # Verificar se o executável foi criado
            exe_path = Path("dist") / "Consolidador_Reports.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✓ Executável criado: {exe_path}")
                print(f"✓ Tamanho: {size_mb:.1f} MB")
                
                # Copiar para o diretório raiz para facilitar acesso
                shutil.copy2(exe_path, "Consolidador_Reports.exe")
                print("✓ Executável copiado para o diretório raiz")
                
                print("\n=== INSTRUÇÕES DE USO ===")
                print("1. Execute o arquivo 'Consolidador_Reports.exe'")
                print("2. Selecione os 3 arquivos Excel das empresas")
                print("3. Aguarde o processamento")
                print("4. O arquivo consolidado será aberto automaticamente")
                print("\n✓ Pronto para uso!")
                
                return True
            else:
                print("✗ Executável não foi criado!")
                return False
        else:
            print("✗ Erro na compilação!")
            print("Erro:", result.stderr)
            return False
            
    except Exception as e:
        print(f"✗ Erro durante a compilação: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if not success:
        print("\nPressione Enter para sair...")
        input()
    else:
        print("\nPressione Enter para sair...")
        input() 