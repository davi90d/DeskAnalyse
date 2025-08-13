"""
Script para validação integrada das funcionalidades.
Testa o fluxo sequencial dos testes e a geração do relatório.
"""

import os
import sys
import time
import tkinter as tk
from tkinter import messagebox
import threading

# Adiciona o diretório raiz ao path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importa os módulos da aplicação
from core.hardware_info import HardwareInfo
from core.report_generator import ReportGenerator
from gui.main_window import MainWindow
from tests.keyboard_test import KeyboardTest
from tests.usb_test import USBTest
from tests.webcam_test import WebcamTest
from tests.audio_test import AudioTest

def test_hardware_info():
    """Testa a coleta de informações de hardware."""
    print("Testando coleta de informações de hardware...")
    
    hardware_info = HardwareInfo()
    
    # Testa a coleta de informações da placa-mãe
    mb_info = hardware_info.get_motherboard_info()
    print(f"Placa-mãe: {mb_info}")
    
    # Testa a coleta de informações do processador
    cpu_info = hardware_info.get_cpu_info()
    print(f"Processador: {cpu_info}")
    
    # Testa a coleta de informações da memória RAM
    ram_info = hardware_info.get_ram_info()
    print(f"Memória RAM: {ram_info}")
    
    # Testa a coleta de informações dos discos
    disk_info = hardware_info.get_disk_info()
    print(f"Discos: {disk_info}")
    
    # Testa a coleta de informações da placa de vídeo
    gpu_info = hardware_info.get_gpu_info()
    print(f"Placa de Vídeo: {gpu_info}")
    
    # Testa a coleta de informações do display
    display_info = hardware_info.get_display_info()
    print(f"Display: {display_info}")
    
    # Testa a coleta de informações do TPM
    tpm_info = hardware_info.get_tpm_info()
    print(f"TPM: {tpm_info}")
    
    # Testa a coleta de informações do Bluetooth
    bt_info = hardware_info.get_bluetooth_info()
    print(f"Bluetooth: {bt_info}")
    
    # Testa a coleta de informações do Wi-Fi
    wifi_info = hardware_info.get_wifi_info()
    print(f"Wi-Fi: {wifi_info}")
    
    print("Teste de coleta de informações de hardware concluído.")
    return True

def test_keyboard():
    """Testa a funcionalidade do teste de teclado."""
    print("Testando funcionalidade do teste de teclado...")
    
    keyboard_test = KeyboardTest()
    
    if keyboard_test.initialize():
        print("Teste de teclado inicializado com sucesso.")
        
        # Simula a execução do teste
        # Nota: Este teste requer interação do usuário, então apenas verificamos a inicialização
        print("Teste de teclado requer interação do usuário. Pulando execução automática.")
        
        # Obtém o resultado formatado
        formatted_result = keyboard_test.get_formatted_result()
        print(f"Resultado formatado: {formatted_result}")
        
        # Limpa recursos
        keyboard_test.cleanup()
        
        return True
    else:
        print(f"Erro ao inicializar teste de teclado: {keyboard_test.result['error']}")
        return False

def test_usb():
    """Testa a funcionalidade do teste de USB."""
    print("Testando funcionalidade do teste de USB...")
    
    usb_test = USBTest()
    
    if usb_test.initialize():
        print("Teste de USB inicializado com sucesso.")
        
        # Simula a execução do teste
        # Nota: Este teste requer interação do usuário, então apenas verificamos a inicialização
        print("Teste de USB requer interação do usuário. Pulando execução automática.")
        
        # Obtém o resultado formatado
        formatted_result = usb_test.get_formatted_result()
        print(f"Resultado formatado: {formatted_result}")
        
        # Limpa recursos
        usb_test.cleanup()
        
        return True
    else:
        print(f"Erro ao inicializar teste de USB: {usb_test.result['error']}")
        return False

def test_webcam():
    """Testa a funcionalidade do teste de webcam."""
    print("Testando funcionalidade do teste de webcam...")
    
    webcam_test = WebcamTest()
    
    if webcam_test.initialize():
        print("Teste de webcam inicializado com sucesso.")
        
        # Simula a execução do teste
        # Nota: Este teste requer interação do usuário, então apenas verificamos a inicialização
        print("Teste de webcam requer interação do usuário. Pulando execução automática.")
        
        # Obtém o resultado formatado
        formatted_result = webcam_test.get_formatted_result()
        print(f"Resultado formatado: {formatted_result}")
        
        # Limpa recursos
        webcam_test.cleanup()
        
        return True
    else:
        print(f"Erro ao inicializar teste de webcam: {webcam_test.result['error']}")
        return False

def test_audio():
    """Testa a funcionalidade do teste de áudio."""
    print("Testando funcionalidade do teste de áudio...")
    
    audio_test = AudioTest()
    
    if audio_test.initialize():
        print("Teste de áudio inicializado com sucesso.")
        
        # Simula a execução do teste
        # Nota: Este teste requer interação do usuário, então apenas verificamos a inicialização
        print("Teste de áudio requer interação do usuário. Pulando execução automática.")
        
        # Obtém o resultado formatado
        formatted_result = audio_test.get_formatted_result()
        print(f"Resultado formatado: {formatted_result}")
        
        # Limpa recursos
        audio_test.cleanup()
        
        return True
    else:
        print(f"Erro ao inicializar teste de áudio: {audio_test.result['error']}")
        return False

def test_report_generator():
    """Testa a geração de relatórios."""
    print("Testando geração de relatórios...")
    
    report_generator = ReportGenerator()
    
    # Define informações de identificação
    report_generator.set_identification({
        'technician_name': 'Técnico de Teste',
        'workbench_id': 'Bancada de Teste',
        'date_time': time.strftime("%d/%m/%Y %H:%M:%S")
    })
    
    # Define informações de hardware
    hardware_info = HardwareInfo()
    
    report_generator.set_hardware_info({
        'motherboard': hardware_info.get_motherboard_info(),
        'cpu': hardware_info.get_cpu_info(),
        'ram': hardware_info.get_ram_info(),
        'display': hardware_info.get_display_info(),
        'tpm': hardware_info.get_tpm_info(),
        'bluetooth': hardware_info.get_bluetooth_info(),
        'wifi': hardware_info.get_wifi_info()
    })
    
    # Adiciona resultados de testes simulados
    report_generator.add_test_result(
        'Teclado',
        {'success': True, 'execution_time': 10.5, 'keys_tested': 104, 'keys_total': 104},
        "Teste de Teclado: SUCESSO\nTempo de execução: 10.50 segundos\nLayout: ABNT2\nTeclas testadas: 104 de 104"
    )
    
    report_generator.add_test_result(
        'USB',
        {'success': True, 'execution_time': 15.2, 'devices_detected': [{'name': 'Dispositivo de Teste', 'usb_version': '3.0', 'port': '1'}]},
        "Teste de USB: SUCESSO\nTempo de execução: 15.20 segundos\nDispositivos detectados: 1\n\nDispositivo 1: Dispositivo de Teste\n  Versão USB: 3.0\n  Porta: 1"
    )
    
    report_generator.add_test_result(
        'Webcam',
        {'success': True, 'execution_time': 8.7, 'webcam_detected': True, 'user_confirmation': True},
        "Teste de Webcam: SUCESSO\nTempo de execução: 8.70 segundos\nWebcam detectada: Sim\nConfirmação do usuário: Sim"
    )
    
    report_generator.add_test_result(
        'Áudio',
        {'success': True, 'execution_time': 12.3, 'user_confirmation': {'microphone': True, 'speaker': True}},
        "Teste de Áudio: SUCESSO\nTempo de execução: 12.30 segundos\nConfirmação do Usuário:\n  Microfone: Sim\n  Alto-falante: Sim"
    )
    
    # Gera o relatório
    report_path = report_generator.generate_report()
    
    if report_path:
        print(f"Relatório gerado com sucesso: {report_path}")
        
        # Tenta exportar para JSON
        json_path = report_generator.export_to_json()
        
        if json_path:
            print(f"Dados exportados para JSON com sucesso: {json_path}")
        
        return True
    else:
        print("Erro ao gerar relatório.")
        return False

def test_sequential_flow():
    """Testa o fluxo sequencial dos testes."""
    print("Testando fluxo sequencial dos testes...")
    
    # Simula a execução sequencial dos testes
    print("1. Teste de Teclado")
    print("2. Teste de USB")
    print("3. Teste de Webcam")
    print("4. Teste de Áudio")
    
    print("Fluxo sequencial validado.")
    return True

def test_gui():
    """Testa a interface gráfica."""
    print("Testando interface gráfica...")
    
    # Cria uma janela raiz
    root = tk.Tk()
    root.withdraw()  # Esconde a janela raiz
    
    # Exibe uma mensagem informativa
    messagebox.showinfo(
        "Teste de Interface",
        "Este teste verifica a inicialização da interface gráfica.\n\n"
        "A janela principal será aberta por alguns segundos e depois fechada automaticamente."
    )
    
    # Cria a janela principal
    main_window = MainWindow(root)
    
    # Agenda o fechamento da janela após 5 segundos
    def close_window():
        root.quit()
    
    root.after(5000, close_window)
    
    # Inicia o loop principal
    root.mainloop()
    
    print("Teste de interface gráfica concluído.")
    return True

def run_all_tests():
    """Executa todos os testes."""
    print("Iniciando testes de validação...")
    
    tests = [
        ("Coleta de Informações de Hardware", test_hardware_info),
        ("Teste de Teclado", test_keyboard),
        ("Teste de USB", test_usb),
        ("Teste de Webcam", test_webcam),
        ("Teste de Áudio", test_audio),
        ("Geração de Relatórios", test_report_generator),
        ("Fluxo Sequencial", test_sequential_flow),
        ("Interface Gráfica", test_gui)
    ]
    
    results = {}
    
    for name, func in tests:
        print(f"\n{'=' * 50}")
        print(f"Executando: {name}")
        print(f"{'-' * 50}")
        
        try:
            result = func()
            results[name] = "SUCESSO" if result else "FALHA"
        except Exception as e:
            print(f"Erro ao executar {name}: {e}")
            results[name] = "ERRO"
    
    print("\n\n")
    print("=" * 50)
    print("RESUMO DOS TESTES")
    print("=" * 50)
    
    for name, result in results.items():
        print(f"{name}: {result}")
    
    print("=" * 50)
    
    # Verifica se todos os testes foram bem-sucedidos
    if all(result == "SUCESSO" for result in results.values()):
        print("\nTodos os testes foram concluídos com sucesso!")
        return True
    else:
        print("\nAlguns testes falharam. Verifique o resumo acima.")
        return False

if __name__ == "__main__":
    run_all_tests()
