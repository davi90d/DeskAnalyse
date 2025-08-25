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

import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

import socket

def main():
    logging.debug("main() function called.")
    
    # Tenta criar um socket para garantir uma única instância
    try:
        # Cria um socket TCP/IP
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # Tenta vincular o socket a um endereço e porta conhecidos
        # Usamos localhost e uma porta arbitrária que provavelmente não estará em uso
        s.bind(("127.0.0.1", 12345)) 
        # Adiciona o socket ao objeto global para evitar que seja coletado pelo garbage collector
        # e para manter a porta ocupada enquanto a aplicação estiver em execução
        global _single_instance_socket
        _single_instance_socket = s
        logging.info("Aplicação iniciada como instância única.")
    except socket.error:
        logging.warning("Outra instância da aplicação já está em execução. Saindo.")
        sys.exit(0) # Sai se outra instância já estiver rodando


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
    

    
    # Cria a janela principal da aplicação
    app = MainWindow(root)
    
    # Inicia o loop principal
    root.mainloop()


# Ponto de entrada principal - garante que o script só é executado diretamente
if __name__ == "__main__":
    # Evita que o script seja executado mais de uma vez
    main()
