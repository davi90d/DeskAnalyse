"""
Módulo para teste de USB.
Implementa teste real de dispositivos USB com seleção e medição de velocidade.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import random
import shutil
import tempfile
import subprocess
import re
import platform

class USBTest:
    """Classe para teste de USB."""
    
    def __init__(self):
        """Inicializa o teste de USB."""
        self.window = None
        self.devices = []
        self.selected_device = None
        self.test_file_size_mb = 100  # Tamanho do arquivo de teste em MB
        self.is_running = False
        self.is_completed = False
        self.result = {
            'success': False,
            'message': '',
            'details': {},
            'error': None
        }
        self.temp_dir = None
        self.test_file = None
    
    def initialize(self):
        """Inicializa o teste de USB."""
        try:
            # Detecta dispositivos USB
            self.devices = self._detect_usb_devices()
            
            if not self.devices:
                self.result['error'] = "Nenhum dispositivo USB detectado"
                return False
            
            self.is_running = True
            return True
        except Exception as e:
            self.result['error'] = str(e)
            return False
    
    def execute(self):
        """Executa o teste de USB."""
        try:
            # Cria a janela de teste
            self.window = tk.Toplevel()
            self.window.title("Teste de USB")
            self.window.geometry("600x500")
            self.window.resizable(True, True)
            self.window.protocol("WM_DELETE_WINDOW", self._on_close)
            
            # Centraliza a janela
            self.window.update_idletasks()
            width = self.window.winfo_width()
            height = self.window.winfo_height()
            x = (self.window.winfo_screenwidth() // 2) - (width // 2)
            y = (self.window.winfo_screenheight() // 2) - (height // 2)
            self.window.geometry(f"{width}x{height}+{x}+{y}")
            
            # Frame principal
            main_frame = ttk.Frame(self.window, padding=10)
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Título
            title_label = ttk.Label(
                main_frame,
                text="Teste de USB - Selecione um dispositivo para testar",
                font=("Arial", 14, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            # Frame para lista de dispositivos
            devices_frame = ttk.LabelFrame(main_frame, text="Dispositivos Detectados", padding=10)
            devices_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # Lista de dispositivos
            self.device_listbox = tk.Listbox(devices_frame, height=10, selectmode=tk.SINGLE)
            self.device_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Scrollbar para a lista
            scrollbar = ttk.Scrollbar(devices_frame, orient=tk.VERTICAL, command=self.device_listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.device_listbox.config(yscrollcommand=scrollbar.set)
            
            # Preenche a lista de dispositivos
            for device in self.devices:
                self.device_listbox.insert(tk.END, f"{device['name']} ({device['path']})")
            
            # Seleciona o primeiro dispositivo por padrão
            if self.devices:
                self.device_listbox.selection_set(0)
            
            # Frame para informações do teste
            info_frame = ttk.LabelFrame(main_frame, text="Informações do Teste", padding=10)
            info_frame.pack(fill=tk.X, pady=10)
            
            # Tamanho do arquivo de teste
            ttk.Label(info_frame, text="Tamanho do arquivo de teste:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            
            self.size_var = tk.StringVar(value=str(self.test_file_size_mb))
            size_entry = ttk.Entry(info_frame, textvariable=self.size_var, width=10)
            size_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
            
            ttk.Label(info_frame, text="MB").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
            
            # Frame para resultados
            self.results_frame = ttk.LabelFrame(main_frame, text="Resultados", padding=10)
            self.results_frame.pack(fill=tk.X, pady=10)
            
            # Velocidade de transferência
            ttk.Label(self.results_frame, text="Velocidade de transferência:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            
            self.speed_var = tk.StringVar(value="Não testado")
            ttk.Label(self.results_frame, textvariable=self.speed_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
            
            # Tipo de USB
            ttk.Label(self.results_frame, text="Tipo de USB:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
            
            self.type_var = tk.StringVar(value="Não detectado")
            ttk.Label(self.results_frame, textvariable=self.type_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
            
            # Barra de progresso
            self.progress_var = tk.DoubleVar(value=0)
            self.progress_bar = ttk.Progressbar(
                main_frame,
                orient=tk.HORIZONTAL,
                length=100,
                mode='determinate',
                variable=self.progress_var
            )
            self.progress_bar.pack(fill=tk.X, pady=10)
            
            # Status
            self.status_var = tk.StringVar(value="Selecione um dispositivo e clique em Iniciar Teste")
            status_label = ttk.Label(main_frame, textvariable=self.status_var)
            status_label.pack(fill=tk.X, pady=5)
            
            # Frame para botões de ação
            action_frame = ttk.Frame(main_frame)
            action_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Botão para iniciar o teste
            self.start_button = ttk.Button(
                action_frame,
                text="Iniciar Teste",
                command=self._start_test
            )
            self.start_button.pack(side=tk.LEFT, padx=5)
            
            # Botão para pular o teste
            skip_button = ttk.Button(
                action_frame,
                text="Pular Teste",
                command=self._skip_test
            )
            skip_button.pack(side=tk.LEFT, padx=5)
            
            # Botão para concluir o teste
            self.complete_button = ttk.Button(
                action_frame,
                text="Concluir Teste",
                state=tk.DISABLED,
                command=self._complete_test
            )
            self.complete_button.pack(side=tk.RIGHT, padx=5)
            
            # Aguarda a conclusão do teste
            self.window.wait_window()
            
            # Retorna o resultado
            return self.is_completed
        except Exception as e:
            self.result['error'] = str(e)
            return False
        finally:
            # Limpa os arquivos temporários
            self._cleanup_temp_files()
    
    def _detect_usb_devices(self):
        """Detecta dispositivos USB conectados."""
        devices = []
        
        try:
            if platform.system() == 'Windows':
                # Detecta dispositivos no Windows
                drives = []
                
                # Obtém as letras de unidade
                output = subprocess.check_output(
                    ["wmic", "logicaldisk", "get", "deviceid,volumename,description"],
                    universal_newlines=True,
                    timeout=10
                )
                
                for line in output.strip().split('\n')[1:]:
                    if not line.strip():
                        continue
                    
                    parts = line.split()
                    if len(parts) >= 2:
                        drive_letter = parts[0]
                        description = ' '.join(parts[1:])
                        
                        # Verifica se é um dispositivo removível
                        if "Removable" in description:
                            name = f"Dispositivo Removível ({drive_letter})"
                            drives.append({
                                'name': name,
                                'path': drive_letter,
                                'type': 'removable'
                            })
                        # Verifica se é um disco fixo
                        elif "Fixed" in description:
                            name = f"Disco Local ({drive_letter})"
                            drives.append({
                                'name': name,
                                'path': drive_letter,
                                'type': 'fixed'
                            })
                
                return drives
            else:
                # Simulação para outros sistemas operacionais
                return [{
                    'name': 'Dispositivo USB Simulado',
                    'path': '/tmp',
                    'type': 'removable'
                }]
        except Exception as e:
            print(f"Erro ao detectar dispositivos USB: {e}")
            # Retorna uma lista vazia em caso de erro
            return []
    
    def _start_test(self):
        """Inicia o teste de USB."""
        # Verifica se um dispositivo foi selecionado
        selection = self.device_listbox.curselection()
        if not selection:
            messagebox.showwarning(
                "Nenhum Dispositivo Selecionado",
                "Por favor, selecione um dispositivo para testar."
            )
            return
        
        # Obtém o dispositivo selecionado
        index = selection[0]
        self.selected_device = self.devices[index]
        
        # Verifica o tamanho do arquivo de teste
        try:
            size_mb = int(self.size_var.get())
            if size_mb <= 0:
                raise ValueError("O tamanho deve ser maior que zero")
            
            self.test_file_size_mb = size_mb
        except ValueError:
            messagebox.showwarning(
                "Tamanho Inválido",
                "Por favor, informe um tamanho válido para o arquivo de teste."
            )
            return
        
        # Desabilita os controles durante o teste
        self.device_listbox.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        # Inicia o teste em uma thread separada
        threading.Thread(target=self._run_test, daemon=True).start()
    
    def _run_test(self):
        """Executa o teste de USB."""
        try:
            # Atualiza o status
            self.status_var.set("Preparando o teste...")
            self.progress_var.set(0)
            
            # Cria um diretório temporário
            self.temp_dir = tempfile.mkdtemp()
            
            # Cria um arquivo de teste
            self.test_file = os.path.join(self.temp_dir, "test_file.bin")
            
            # Atualiza o status
            self.status_var.set("Criando arquivo de teste...")
            
            # Cria o arquivo de teste
            with open(self.test_file, 'wb') as f:
                # Escreve em blocos de 1MB
                block_size = 1024 * 1024
                for i in range(self.test_file_size_mb):
                    # Gera dados aleatórios
                    data = os.urandom(block_size)
                    f.write(data)
                    
                    # Atualiza o progresso
                    progress = (i + 1) / self.test_file_size_mb * 50
                    self.window.after(0, lambda p=progress: self.progress_var.set(p))
                    
                    # Atualiza o status
                    self.window.after(0, lambda i=i, t=self.test_file_size_mb: self.status_var.set(
                        f"Criando arquivo de teste... {i+1}/{t} MB"
                    ))
            
            # Atualiza o status
            self.window.after(0, lambda: self.status_var.set("Iniciando transferência para o dispositivo..."))
            
            # Define o caminho de destino
            if platform.system() == 'Windows':
                dest_path = os.path.join(self.selected_device['path'], "test_file.bin")
            else:
                dest_path = os.path.join(self.selected_device['path'], "test_file.bin")
            
            # Mede o tempo de transferência
            start_time = time.time()
            
            # Copia o arquivo para o dispositivo
            shutil.copy2(self.test_file, dest_path)
            
            # Calcula o tempo de transferência
            end_time = time.time()
            transfer_time = end_time - start_time
            
            # Calcula a velocidade de transferência em MB/s
            speed_mbps = self.test_file_size_mb / transfer_time
            
            # Atualiza o progresso
            self.window.after(0, lambda: self.progress_var.set(100))
            
            # Atualiza o status
            self.window.after(0, lambda: self.status_var.set("Teste concluído com sucesso!"))
            
            # Atualiza os resultados
            self.window.after(0, lambda: self.speed_var.set(f"{speed_mbps:.2f} MB/s"))
            
            # Determina o tipo de USB com base na velocidade
            usb_type = self._determine_usb_type(speed_mbps)
            self.window.after(0, lambda: self.type_var.set(usb_type))
            
            # Habilita o botão de concluir
            self.window.after(0, lambda: self.complete_button.config(state=tk.NORMAL))
            
            # Armazena os resultados
            self.result['success'] = True
            self.result['message'] = "Teste concluído com sucesso"
            self.result['details'] = {
                'device': self.selected_device['name'],
                'speed_mbps': speed_mbps,
                'usb_type': usb_type
            }
            
            # Remove o arquivo de teste do dispositivo
            try:
                os.remove(dest_path)
            except:
                pass
        except Exception as e:
            # Atualiza o status
            self.window.after(0, lambda: self.status_var.set(f"Erro: {e}"))
            
            # Armazena o erro
            self.result['error'] = str(e)
            
            # Habilita o botão de iniciar
            self.window.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            self.window.after(0, lambda: self.device_listbox.config(state=tk.NORMAL))
        finally:
            # Limpa os arquivos temporários
            self._cleanup_temp_files()
    
    def _determine_usb_type(self, speed_mbps):
        """Determina o tipo de USB com base na velocidade de transferência."""
        if speed_mbps < 30:
            return "USB 2.0"
        elif speed_mbps < 100:
            return "USB 3.2 Gen 1 (5 Gbps)"
        else:
            return "USB 3.2 Gen 2 (10 Gbps)"
    
    def _cleanup_temp_files(self):
        """Limpa os arquivos temporários."""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
        except:
            pass
    
    def _skip_test(self):
        """Pula o teste."""
        if messagebox.askyesno(
            "Pular Teste",
            "Tem certeza que deseja pular o teste de USB?\n\n"
            "O teste será marcado como não concluído."
        ):
            self.result['success'] = False
            self.result['message'] = "Teste pulado pelo usuário"
            self.is_completed = False
            self.window.destroy()
    
    def _complete_test(self):
        """Conclui o teste."""
        self.is_completed = True
        self.window.destroy()
    
    def _on_close(self):
        """Manipulador de evento de fechamento da janela."""
        if messagebox.askyesno(
            "Fechar Teste",
            "Tem certeza que deseja fechar o teste de USB?\n\n"
            "O teste será marcado como não concluído."
        ):
            self.result['success'] = False
            self.result['message'] = "Teste interrompido pelo usuário"
            self.is_completed = False
            self.window.destroy()
    
    def get_result(self):
        """Retorna o resultado do teste."""
        return self.result
    
    def get_formatted_result(self):
        """Retorna o resultado formatado do teste."""
        if self.result['success']:
            return f"Teste de USB: SUCESSO\n" \
                   f"Dispositivo: {self.result['details'].get('device', 'Não disponível')}\n" \
                   f"Velocidade: {self.result['details'].get('speed_mbps', 0):.2f} MB/s\n" \
                   f"Tipo de USB: {self.result['details'].get('usb_type', 'Não detectado')}"
        else:
            if self.result['error']:
                return f"Teste de USB: FALHA\n" \
                       f"Erro: {self.result['error']}"
            else:
                return f"Teste de USB: FALHA\n" \
                       f"Motivo: {self.result['message']}"
