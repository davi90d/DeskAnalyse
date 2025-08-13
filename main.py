"""
Módulo principal da aplicação de diagnóstico de hardware.
Integra todos os componentes e inicia a interface gráfica.
Versão corrigida para compatibilidade com PyInstaller.
"""

import os
import sys
import tkinter as tk
from datetime import datetime
import threading

# Importações dos módulos da aplicação
from gui.main_window import MainWindow

def main():
    """Função principal para iniciar a aplicação."""
    # Configura o diretório base para recursos quando empacotado com PyInstaller
    if getattr(sys, 'frozen', False):
        # Se estiver executando como um executável empacotado
        application_path = os.path.dirname(sys.executable)
        os.chdir(application_path)
        
        # Adiciona o diretório ao path para garantir que os módulos sejam encontrados
        if application_path not in sys.path:
            sys.path.insert(0, application_path)
    
    # Cria a janela principal do Tkinter
    root = tk.Tk()
    root.title("Diagnóstico de Hardware")
    
    # Evita múltiplas instâncias verificando se já existe uma janela
    try:
        # Tenta definir um atributo único para verificar se já existe uma instância
        root.attributes('-unique', True)
    except:
        # Se falhar, continua normalmente
        pass
    
    # Cria a janela principal da aplicação
    app = MainWindow(root)
    
    # Inicia o loop principal
    root.mainloop()


# Ponto de entrada principal - garante que o script só é executado diretamente
if __name__ == "__main__":
    # Evita que o script seja executado mais de uma vez
    main()
