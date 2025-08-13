"""
Módulo para gerenciamento da interface gráfica principal.
Implementa a janela principal da aplicação com abas reorganizadas.
Versão corrigida para compatibilidade com PyInstaller e novos requisitos.
"""

import os
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
from datetime import datetime
import queue

# Importa os módulos da aplicação
from core.hardware_info import HardwareInfo
from core.report_generator import ReportGenerator
from tests.keyboard_test import KeyboardTest
from tests.usb_test import USBTest
from tests.webcam_test import WebcamTest
from tests.audio_test import AudioTest

class MainWindow:
    """Classe para gerenciamento da janela principal da aplicação."""
    
    def __init__(self, root):
        """Inicializa a janela principal."""
        self.root = root
        self.root.title("Diagnóstico de Hardware")
        self.root.geometry("900x650")  
        self.root.minsize(900, 650)  
        
        
        # Flag para controlar se a coleta de hardware já foi iniciada
        self.hardware_collection_started = False
        self.hardware_collection_running = False
        
        # Flag para controlar se o cadastro foi concluído
        self.registration_completed = False
        
        # Verifica privilégios de administrador
        self._check_admin_privileges()
        
        # Inicializa variáveis
        self.hardware_info = HardwareInfo()
        self.report_generator = ReportGenerator()
        self.technician_name = tk.StringVar()
        self.workbench_id = tk.StringVar()
        self.test_queue = queue.Queue()
        self.selected_tests = {
            "keyboard": tk.BooleanVar(value=False),
            "usb": tk.BooleanVar(value=False),
            "webcam": tk.BooleanVar(value=False),
            "audio": tk.BooleanVar(value=False)
        }
        
        # Flag para controlar se os testes estão em execução
        self.tests_running = False
        
        # Configura o estilo
        self._setup_styles()
        
        # Cria a interface
        self._create_interface()
        
        # Agenda a coleta de informações de hardware após a interface ser criada
        # Isso evita problemas de threading durante a inicialização
        self.root.after(100, self._collect_hardware_info)
    
    def _check_admin_privileges(self):
        """Verifica se a aplicação está sendo executada como administrador."""
        try:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                messagebox.showerror(
                    "Erro de Privilégios",
                    "Esta aplicação requer privilégios de administrador.\n\n"
                    "Por favor, feche a aplicação e execute-a como administrador."
                )
                self.root.after(3000, self.root.destroy)
        except Exception as e:
            messagebox.showwarning(
                "Aviso",
                f"Não foi possível verificar privilégios de administrador: {e}\n\n"
                "A aplicação pode não funcionar corretamente sem privilégios de administrador."
            )
    
    def _setup_styles(self):
        """Configura os estilos da interface."""
        style = ttk.Style()
        
        # Estilos para abas
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", padding=[10, 5], font=("Arial", 10))
        
        # Estilos para frames
        style.configure("Card.TFrame", relief="raised", borderwidth=1)
        
        # Estilos para labels
        style.configure("Title.TLabel", font=("Arial", 14, "bold"))
        style.configure("Subtitle.TLabel", font=("Arial", 12, "bold"))
        style.configure("Info.TLabel", font=("Arial", 10))
        style.configure("Value.TLabel", font=("Arial", 10, "bold"))
        
        # Estilos para botões
        style.configure("Primary.TButton", font=("Arial", 10, "bold"))
        style.configure("Secondary.TButton", font=("Arial", 10))
        style.configure("Success.TButton", font=("Arial", 10, "bold"), background="#4CAF50")
        style.configure("Warning.TButton", font=("Arial", 10, "bold"), background="#FFC107")
    
    def _create_interface(self):
        """Cria a interface gráfica."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            title_frame,
            text="Diagnóstico de Hardware",
            style="Title.TLabel"
        ).pack(side=tk.LEFT)
        
        # Botão para atualizar informações de hardware
        self.refresh_button = ttk.Button(
            title_frame,
            text="Atualizar Informações",
            style="Secondary.TButton",
            command=self._refresh_hardware_info
        )
        self.refresh_button.pack(side=tk.RIGHT, padx=5)
        
        # Frame para identificação
        id_frame = ttk.LabelFrame(main_frame, text="Identificação", padding=10)
        id_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Campos de identificação
        ttk.Label(id_frame, text="Nome do Técnico:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.technician_entry = ttk.Entry(id_frame, textvariable=self.technician_name, width=30)
        self.technician_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(id_frame, text="ID da Bancada:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.workbench_entry = ttk.Entry(id_frame, textvariable=self.workbench_id, width=30)
        self.workbench_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Botão para concluir cadastro
        self.complete_registration_button = ttk.Button(
            id_frame,
            text="Concluir Cadastro",
            style="Success.TButton",
            command=self._complete_registration
        )
        self.complete_registration_button.grid(row=1, column=2, padx=5, pady=5)
        
        # Notebook (abas)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Aba de Informações de Hardware
        self.hw_info_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.hw_info_tab, text="Informações de Hardware")
        
        # Aba de Testes
        self.tests_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tests_tab, text="Testes")
        
        # Cria o conteúdo das abas
        self._create_hardware_info_tab()
        self._create_tests_tab()
        
        # Frame para botões de ação
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Botões de ação - Usando grid para garantir visibilidade
        self.run_tests_button = ttk.Button(
            action_frame,
            text="Executar Testes Selecionados",
            style="Primary.TButton",
            command=self._run_selected_tests,
            state=tk.DISABLED
        )
        self.run_tests_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.run_all_button = ttk.Button(
            action_frame,
            text="Teste Completo",
            style="Primary.TButton",
            command=self._run_all_tests,
            state=tk.DISABLED
        )
        self.run_all_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.report_button = ttk.Button(
            action_frame,
            text="Gerar Relatório",
            style="Secondary.TButton",
            command=self._generate_report
        )
        self.report_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        # Garante que as colunas ocupem espaço proporcional
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)
        action_frame.columnconfigure(2, weight=1)
        
        # Barra de status
        self.status_var = tk.StringVar(value="Inicializando...")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def _create_hardware_info_tab(self):
        """Cria o conteúdo da aba de Informações de Hardware."""
        # Frame para informações da placa-mãe
        mb_frame = ttk.LabelFrame(self.hw_info_tab, text="Placa-Mãe", padding=10)
        mb_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.mb_manufacturer_var = tk.StringVar(value="Carregando...")
        self.mb_model_var = tk.StringVar(value="Carregando...")
        self.mb_serial_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(mb_frame, text="Fabricante:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(mb_frame, textvariable=self.mb_manufacturer_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(mb_frame, text="Modelo:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(mb_frame, textvariable=self.mb_model_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(mb_frame, text="Número de Série:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(mb_frame, textvariable=self.mb_serial_var, style="Value.TLabel").grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do processador
        cpu_frame = ttk.LabelFrame(self.hw_info_tab, text="Processador", padding=10)
        cpu_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        self.cpu_brand_var = tk.StringVar(value="Carregando...")
        self.cpu_model_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(cpu_frame, text="Marca:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(cpu_frame, textvariable=self.cpu_brand_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(cpu_frame, text="Modelo:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(cpu_frame, textvariable=self.cpu_model_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações da memória RAM
        ram_frame = ttk.LabelFrame(self.hw_info_tab, text="Memória RAM", padding=10)
        ram_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        self.ram_total_var = tk.StringVar(value="Carregando...")
        self.ram_slots_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(ram_frame, text="Total:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(ram_frame, textvariable=self.ram_total_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(ram_frame, text="Slots Usados:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(ram_frame, textvariable=self.ram_slots_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para detalhes dos módulos de RAM
        self.ram_modules_frame = ttk.Frame(ram_frame)
        self.ram_modules_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        # Frame para informações dos discos
        disk_frame = ttk.LabelFrame(self.hw_info_tab, text="Discos", padding=10)
        disk_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        # Treeview para discos
        self.disk_tree = ttk.Treeview(disk_frame, columns=("model", "type", "size", "free"), show="headings", height=4)
        self.disk_tree.heading("model", text="Modelo")
        self.disk_tree.heading("type", text="Tipo")
        self.disk_tree.heading("size", text="Tamanho")
        self.disk_tree.heading("free", text="Livre")
        
        self.disk_tree.column("model", width=150)
        self.disk_tree.column("type", width=80)
        self.disk_tree.column("size", width=80)
        self.disk_tree.column("free", width=80)
        
        self.disk_tree.pack(fill=tk.BOTH, expand=True)
        
        # Frame para informações da GPU
        gpu_frame = ttk.LabelFrame(self.hw_info_tab, text="Placa de Vídeo", padding=10)
        gpu_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        
        # Treeview para GPUs
        self.gpu_tree = ttk.Treeview(gpu_frame, columns=("name"), show="headings", height=2)
        self.gpu_tree.heading("name", text="Modelo")
        self.gpu_tree.column("name", width=300)
        self.gpu_tree.pack(fill=tk.BOTH, expand=True)
        
        # Frame para informações do display
        display_frame = ttk.LabelFrame(self.hw_info_tab, text="Display", padding=10)
        display_frame.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
        
        self.display_resolution_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(display_frame, text="Resolução:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(display_frame, textvariable=self.display_resolution_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do TPM
        tpm_frame = ttk.LabelFrame(self.hw_info_tab, text="TPM", padding=10)
        tpm_frame.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        
        self.tpm_version_var = tk.StringVar(value="Carregando...")
        self.tpm_status_var = tk.StringVar(value="Carregando...")
        self.tpm_manufacturer_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(tpm_frame, text="Versão:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(tpm_frame, textvariable=self.tpm_version_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(tpm_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(tpm_frame, textvariable=self.tpm_status_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(tpm_frame, text="Fabricante:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(tpm_frame, textvariable=self.tpm_manufacturer_var, style="Value.TLabel").grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do Bluetooth
        bt_frame = ttk.LabelFrame(self.hw_info_tab, text="Bluetooth", padding=10)
        bt_frame.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)
        
        self.bt_device_var = tk.StringVar(value="Carregando...")
        self.bt_status_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(bt_frame, text="Dispositivo:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(bt_frame, textvariable=self.bt_device_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(bt_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(bt_frame, textvariable=self.bt_status_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do Wi-Fi
        wifi_frame = ttk.LabelFrame(self.hw_info_tab, text="Wi-Fi", padding=10)
        wifi_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        self.wifi_adapter_var = tk.StringVar(value="Carregando...")
        self.wifi_status_var = tk.StringVar(value="Carregando...")
        self.wifi_ssid_var = tk.StringVar(value="Carregando...")
        
        ttk.Label(wifi_frame, text="Adaptador:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(wifi_frame, textvariable=self.wifi_adapter_var, style="Value.TLabel").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(wifi_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(wifi_frame, textvariable=self.wifi_status_var, style="Value.TLabel").grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(wifi_frame, text="SSID:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(wifi_frame, textvariable=self.wifi_ssid_var, style="Value.TLabel").grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Configura o grid
        for i in range(5):
            self.hw_info_tab.rowconfigure(i, weight=1)
        
        self.hw_info_tab.columnconfigure(0, weight=1)
        self.hw_info_tab.columnconfigure(1, weight=1)
    
    def _create_tests_tab(self):
        """Cria o conteúdo da aba de Testes."""
        # Frame para seleção de testes
        selection_frame = ttk.LabelFrame(self.tests_tab, text="Selecione os Testes", padding=10)
        selection_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Checkboxes para seleção de testes
        ttk.Checkbutton(
            selection_frame,
            text="Teste de Teclado",
            variable=self.selected_tests["keyboard"]
        ).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            selection_frame,
            text="Teste de USB",
            variable=self.selected_tests["usb"]
        ).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            selection_frame,
            text="Teste de Webcam",
            variable=self.selected_tests["webcam"]
        ).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            selection_frame,
            text="Teste de Áudio",
            variable=self.selected_tests["audio"]
        ).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Descrição dos testes
        description_frame = ttk.LabelFrame(self.tests_tab, text="Descrição dos Testes", padding=10)
        description_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        descriptions = {
            "keyboard": "Teste de Teclado: Verifica o funcionamento de todas as teclas do teclado.",
            "usb": "Teste de USB: Verifica a conexão e velocidade de dispositivos USB.",
            "webcam": "Teste de Webcam: Verifica o funcionamento da webcam.",
            "audio": "Teste de Áudio: Verifica o funcionamento do microfone e alto-falantes."
        }
        
        for i, (test_id, description) in enumerate(descriptions.items()):
            ttk.Label(description_frame, text=description).pack(anchor=tk.W, padx=5, pady=2)
    
    def _complete_registration(self):
        """Conclui o cadastro de usuário/bancada."""
        # Verifica se os campos de identificação foram preenchidos
        if not self.technician_name.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "O campo 'Nome do Técnico' é obrigatório.\n\n"
                "Por favor, preencha este campo antes de continuar."
            )
            return
        
        if not self.workbench_id.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "O campo 'ID da Bancada' é obrigatório.\n\n"
                "Por favor, preencha este campo antes de continuar."
            )
            return
        
        # Confirma a conclusão do cadastro
        if messagebox.askyesno(
            "Concluir Cadastro",
            "Deseja concluir o cadastro?\n\n"
            "Os campos de identificação serão bloqueados para edição."
        ):
            # Marca o cadastro como concluído
            self.registration_completed = True
            
            # Desabilita os campos de identificação
            self.technician_entry.config(state=tk.DISABLED)
            self.workbench_entry.config(state=tk.DISABLED)
            self.complete_registration_button.config(state=tk.DISABLED)
            
            # Habilita os botões de teste
            self.run_tests_button.config(state=tk.NORMAL)
            self.run_all_button.config(state=tk.NORMAL)
            
            # Atualiza o status
            self.status_var.set(f"Cadastro concluído. Técnico: {self.technician_name.get()}, Bancada: {self.workbench_id.get()}")
            
            # Exibe mensagem de sucesso
            messagebox.showinfo(
                "Cadastro Concluído",
                "Cadastro concluído com sucesso!\n\n"
                "Agora você pode executar os testes."
            )
    
    def _refresh_hardware_info(self):
        """Atualiza as informações de hardware."""
        # Verifica se já está coletando informações
        if self.hardware_collection_running:
            messagebox.showinfo(
                "Coleta em Andamento",
                "A coleta de informações de hardware já está em andamento.\n\n"
                "Por favor, aguarde a conclusão da coleta atual."
            )
            return
        
        # Desabilita o botão de atualização durante a coleta
        self.refresh_button.config(state=tk.DISABLED)
        
        # Atualiza o status
        self.status_var.set("Atualizando informações de hardware...")
        
        # Limpa as informações atuais
        self._reset_hardware_info_display()
        
        # Coleta as informações de hardware
        self._collect_hardware_info(force=True)
    
    def _reset_hardware_info_display(self):
        """Limpa as informações de hardware exibidas."""
        # Placa-mãe
        self.mb_manufacturer_var.set("Carregando...")
        self.mb_model_var.set("Carregando...")
        self.mb_serial_var.set("Carregando...")
        
        # Processador
        self.cpu_brand_var.set("Carregando...")
        self.cpu_model_var.set("Carregando...")
        
        # Memória RAM
        self.ram_total_var.set("Carregando...")
        self.ram_slots_var.set("Carregando...")
        
        # Limpa o frame de módulos de RAM
        self._clear_frame(self.ram_modules_frame)
        
        # Discos
        self.disk_tree.delete(*self.disk_tree.get_children())
        
        # GPU
        self.gpu_tree.delete(*self.gpu_tree.get_children())
        
        # Display
        self.display_resolution_var.set("Carregando...")
        
        # TPM
        self.tpm_version_var.set("Carregando...")
        self.tpm_status_var.set("Carregando...")
        self.tpm_manufacturer_var.set("Carregando...")
        
        # Bluetooth
        self.bt_device_var.set("Carregando...")
        self.bt_status_var.set("Carregando...")
        
        # Wi-Fi
        self.wifi_adapter_var.set("Carregando...")
        self.wifi_status_var.set("Carregando...")
        self.wifi_ssid_var.set("Carregando...")
    
    def _collect_hardware_info(self, force=False):
        """Coleta informações de hardware."""
        # Evita múltiplas coletas simultâneas
        if self.hardware_collection_running:
            return
        
        # Evita coletas repetidas, a menos que seja forçado
        if self.hardware_collection_started and not force:
            return
        
        self.hardware_collection_started = True
        self.hardware_collection_running = True
        self.status_var.set("Coletando informações de hardware...")
        
        # Desabilita os botões durante a coleta
        self.refresh_button.config(state=tk.DISABLED)
        if not self.registration_completed:
            self.run_tests_button.config(state=tk.DISABLED)
            self.run_all_button.config(state=tk.DISABLED)
        self.report_button.config(state=tk.DISABLED)
        
        # Executa a coleta em uma thread separada
        def collect_info():
            try:
                # Coleta informações da placa-mãe
                mb_info = self.hardware_info.get_motherboard_info()
                self.root.after(0, lambda: self.mb_manufacturer_var.set(mb_info['manufacturer']))
                self.root.after(0, lambda: self.mb_model_var.set(mb_info['model']))
                self.root.after(0, lambda: self.mb_serial_var.set(mb_info['serial_number']))
                
                # Coleta informações do processador
                cpu_info = self.hardware_info.get_cpu_info()
                self.root.after(0, lambda: self.cpu_brand_var.set(cpu_info['brand']))
                self.root.after(0, lambda: self.cpu_model_var.set(cpu_info['model']))
                
                # Coleta informações da memória RAM
                ram_info = self.hardware_info.get_ram_info()
                self.root.after(0, lambda: self.ram_total_var.set(ram_info['total']))
                self.root.after(0, lambda: self.ram_slots_var.set(ram_info['slots_used']))
                
                # Limpa o frame de módulos de RAM
                self.root.after(0, lambda: self._clear_frame(self.ram_modules_frame))
                
                # Adiciona informações dos módulos de RAM
                if 'modules' in ram_info and ram_info['modules']:
                    for i, module in enumerate(ram_info['modules']):
                        module_text = f"Slot {i+1}: {module.get('size', 'N/A')} - {module.get('manufacturer', 'N/A')}"
                        self.root.after(0, lambda t=module_text: ttk.Label(self.ram_modules_frame, text=t).pack(anchor=tk.W))
                
                # Coleta informações dos discos
                disk_info = self.hardware_info.get_disk_info()
                
                # Limpa a treeview de discos
                self.root.after(0, lambda: self.disk_tree.delete(*self.disk_tree.get_children()))
                
                # Adiciona informações dos discos
                for i, disk in enumerate(disk_info):
                    self.root.after(0, lambda d=disk, i=i: self.disk_tree.insert(
                        "",
                        i,
                        values=(
                            d.get('model', 'Não disponível'),
                            d.get('type', 'Não disponível'),
                            d.get('size', 'Não disponível'),
                            d.get('free_space', 'Não disponível')
                        )
                    ))
                
                # Coleta informações da GPU
                gpu_info = self.hardware_info.get_gpu_info()
                
                # Limpa a treeview de GPUs
                self.root.after(0, lambda: self.gpu_tree.delete(*self.gpu_tree.get_children()))
                
                # Adiciona informações das GPUs
                for i, gpu in enumerate(gpu_info):
                    self.root.after(0, lambda g=gpu, i=i: self.gpu_tree.insert(
                        "",
                        i,
                        values=(g.get('name', 'Não disponível'),)
                    ))
                
                # Coleta informações do display
                display_info = self.hardware_info.get_display_info()
                self.root.after(0, lambda: self.display_resolution_var.set(display_info['resolution']))
                
                # Coleta informações do TPM
                tpm_info = self.hardware_info.get_tpm_info()
                self.root.after(0, lambda: self.tpm_version_var.set(tpm_info['version']))
                self.root.after(0, lambda: self.tpm_status_var.set(tpm_info['status']))
                self.root.after(0, lambda: self.tpm_manufacturer_var.set(tpm_info['manufacturer']))
                
                # Coleta informações do Bluetooth
                bt_info = self.hardware_info.get_bluetooth_info()
                self.root.after(0, lambda: self.bt_device_var.set(bt_info['device_name']))
                self.root.after(0, lambda: self.bt_status_var.set(bt_info['device_status']))
                
                # Coleta informações do Wi-Fi
                wifi_info = self.hardware_info.get_wifi_info()
                self.root.after(0, lambda: self.wifi_adapter_var.set(wifi_info['adapter_name']))
                self.root.after(0, lambda: self.wifi_status_var.set(wifi_info['adapter_status']))
                self.root.after(0, lambda: self.wifi_ssid_var.set(wifi_info['connected_ssid']))
                
                # Atualiza o status
                self.root.after(0, lambda: self.status_var.set("Informações de hardware coletadas com sucesso."))
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"Erro ao coletar informações de hardware: {e}"))
                self.root.after(0, lambda: messagebox.showerror(
                    "Erro",
                    f"Ocorreu um erro ao coletar informações de hardware:\n\n{e}"
                ))
            finally:
                # Reabilita os botões
                self.root.after(0, lambda: self.refresh_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.report_button.config(state=tk.NORMAL))
                
                # Habilita os botões de teste apenas se o cadastro estiver concluído
                if self.registration_completed:
                    self.root.after(0, lambda: self.run_tests_button.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.run_all_button.config(state=tk.NORMAL))
                
                # Marca que a coleta terminou
                self.hardware_collection_running = False
        
        # Inicia a thread
        threading.Thread(target=collect_info, daemon=True).start()
    
    def _clear_frame(self, frame):
        """Limpa todos os widgets de um frame."""
        for widget in frame.winfo_children():
            widget.destroy()
    
    def _run_selected_tests(self):
        """Executa os testes selecionados."""
        # Verifica se o cadastro foi concluído
        if not self.registration_completed:
            messagebox.showwarning(
                "Cadastro Incompleto",
                "O cadastro de usuário/bancada não foi concluído.\n\n"
                "Por favor, conclua o cadastro antes de executar os testes."
            )
            return
        
        # Verifica se algum teste foi selecionado
        selected = [test_id for test_id, var in self.selected_tests.items() if var.get()]
        
        if not selected:
            messagebox.showwarning(
                "Nenhum Teste Selecionado",
                "Nenhum teste foi selecionado.\n\n"
                "Por favor, selecione pelo menos um teste para executar."
            )
            return
        
        # Verifica se já há testes em execução
        if self.tests_running:
            messagebox.showwarning(
                "Testes em Execução",
                "Já existem testes em execução.\n\n"
                "Por favor, aguarde a conclusão dos testes atuais."
            )
            return
        
        # Marca que os testes estão em execução
        self.tests_running = True
        
        # Desabilita os botões durante a execução dos testes
        self.run_tests_button.config(state=tk.DISABLED)
        self.run_all_button.config(state=tk.DISABLED)
        
        # Limpa a fila de testes
        while not self.test_queue.empty():
            self.test_queue.get()
        
        # Adiciona os testes selecionados à fila
        for test_id in selected:
            self.test_queue.put(test_id)
        
        # Inicia a execução dos testes
        self._execute_next_test()
    
    def _run_all_tests(self):
        """Executa todos os testes disponíveis."""
        # Verifica se o cadastro foi concluído
        if not self.registration_completed:
            messagebox.showwarning(
                "Cadastro Incompleto",
                "O cadastro de usuário/bancada não foi concluído.\n\n"
                "Por favor, conclua o cadastro antes de executar os testes."
            )
            return
        
        # Verifica se já há testes em execução
        if self.tests_running:
            messagebox.showwarning(
                "Testes em Execução",
                "Já existem testes em execução.\n\n"
                "Por favor, aguarde a conclusão dos testes atuais."
            )
            return
        
        # Marca que os testes estão em execução
        self.tests_running = True
        
        # Desabilita os botões durante a execução dos testes
        self.run_tests_button.config(state=tk.DISABLED)
        self.run_all_button.config(state=tk.DISABLED)
        
        # Limpa a fila de testes
        while not self.test_queue.empty():
            self.test_queue.get()
        
        # Adiciona todos os testes à fila
        for test_id in self.selected_tests.keys():
            self.test_queue.put(test_id)
        
        # Inicia a execução dos testes
        self._execute_next_test()
    
    def _execute_next_test(self):
        """Executa o próximo teste na fila."""
        if self.test_queue.empty():
            # Todos os testes foram concluídos
            self.tests_running = False
            self.run_tests_button.config(state=tk.NORMAL)
            self.run_all_button.config(state=tk.NORMAL)
            self.status_var.set("Todos os testes foram concluídos.")
            
            messagebox.showinfo(
                "Testes Concluídos",
                "Todos os testes foram concluídos.\n\n"
                "Você pode gerar um relatório com os resultados."
            )
            return
        
        # Obtém o próximo teste
        test_id = self.test_queue.get()
        
        # Executa o teste correspondente
        if test_id == "keyboard":
            self._run_keyboard_test()
        elif test_id == "usb":
            self._run_usb_test()
        elif test_id == "webcam":
            self._run_webcam_test()
        elif test_id == "audio":
            self._run_audio_test()
        else:
            # Teste desconhecido, passa para o próximo
            self.root.after(100, self._execute_next_test)
    
    def _run_keyboard_test(self):
        """Executa o teste de teclado."""
        self.status_var.set("Executando teste de teclado...")
        
        # Inicializa o teste
        keyboard_test = KeyboardTest()
        
        if keyboard_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    keyboard_test.execute()
                    
                    # Registra o resultado
                    self.report_generator.add_test_result(
                        "Teclado",
                        keyboard_test.get_result(),
                        keyboard_test.get_formatted_result()
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
                except Exception as e:
                    self.status_var.set(f"Erro ao executar teste de teclado: {e}")
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de teclado:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            self.status_var.set(f"Erro ao inicializar teste de teclado: {keyboard_test.result['error']}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de teclado:\n\n{keyboard_test.result['error']}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)
    
    def _run_usb_test(self):
        """Executa o teste de USB."""
        self.status_var.set("Executando teste de USB...")
        
        # Inicializa o teste
        usb_test = USBTest()
        
        if usb_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    usb_test.execute()
                    
                    # Registra o resultado
                    self.report_generator.add_test_result(
                        "USB",
                        usb_test.get_result(),
                        usb_test.get_formatted_result()
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
                except Exception as e:
                    self.status_var.set(f"Erro ao executar teste de USB: {e}")
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de USB:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            self.status_var.set(f"Erro ao inicializar teste de USB: {usb_test.result['error']}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de USB:\n\n{usb_test.result['error']}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)
    
    def _run_webcam_test(self):
        """Executa o teste de webcam."""
        self.status_var.set("Executando teste de webcam...")
        
        # Inicializa o teste
        webcam_test = WebcamTest()
        
        if webcam_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    webcam_test.execute()
                    
                    # Registra o resultado
                    self.report_generator.add_test_result(
                        "Webcam",
                        webcam_test.get_result(),
                        webcam_test.get_formatted_result()
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
                except Exception as e:
                    self.status_var.set(f"Erro ao executar teste de webcam: {e}")
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de webcam:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            self.status_var.set(f"Erro ao inicializar teste de webcam: {webcam_test.result['error']}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de webcam:\n\n{webcam_test.result['error']}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)
    
    def _run_audio_test(self):
        """Executa o teste de áudio."""
        self.status_var.set("Executando teste de áudio...")
        
        # Inicializa o teste
        audio_test = AudioTest()
        
        if audio_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    audio_test.execute()
                    
                    # Registra o resultado
                    self.report_generator.add_test_result(
                        "Áudio",
                        audio_test.get_result(),
                        audio_test.get_formatted_result()
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
                except Exception as e:
                    self.status_var.set(f"Erro ao executar teste de áudio: {e}")
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de áudio:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            self.status_var.set(f"Erro ao inicializar teste de áudio: {audio_test.result['error']}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de áudio:\n\n{audio_test.result['error']}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)
    
    def _generate_report(self):
        """Gera um relatório com os resultados dos testes."""
        # Verifica se os campos de identificação foram preenchidos
        if not self.technician_name.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "O campo 'Nome do Técnico' é obrigatório.\n\n"
                "Por favor, preencha este campo antes de continuar."
            )
            return
        
        if not self.workbench_id.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "O campo 'ID da Bancada' é obrigatório.\n\n"
                "Por favor, preencha este campo antes de continuar."
            )
            return
        
        # Adiciona informações de hardware ao relatório
        self.report_generator.set_hardware_info({
            'motherboard': {
                'manufacturer': self.mb_manufacturer_var.get(),
                'model': self.mb_model_var.get(),
                'serial_number': self.mb_serial_var.get()
            },
            'cpu': {
                'brand': self.cpu_brand_var.get(),
                'model': self.cpu_model_var.get()
            },
            'ram': {
                'total': self.ram_total_var.get(),
                'slots_used': self.ram_slots_var.get()
            },
            'display': {
                'resolution': self.display_resolution_var.get()
            },
            'tpm': {
                'version': self.tpm_version_var.get(),
                'status': self.tpm_status_var.get(),
                'manufacturer': self.tpm_manufacturer_var.get()
            },
            'bluetooth': {
                'device_name': self.bt_device_var.get(),
                'device_status': self.bt_status_var.get()
            },
            'wifi': {
                'adapter_name': self.wifi_adapter_var.get(),
                'adapter_status': self.wifi_status_var.get(),
                'connected_ssid': self.wifi_ssid_var.get()
            }
        })
        
        # Adiciona informações de identificação ao relatório
        self.report_generator.set_identification({
            'technician_name': self.technician_name.get(),
            'workbench_id': self.workbench_id.get(),
            'date_time': datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        })
        
        # Gera o relatório
        try:
            report_path = self.report_generator.generate_report()
            
            self.status_var.set(f"Relatório gerado com sucesso: {report_path}")
            messagebox.showinfo(
                "Relatório Gerado",
                f"O relatório foi gerado com sucesso:\n\n{report_path}\n\n"
                "Você pode abrir o arquivo para visualizar os resultados."
            )
        except Exception as e:
            self.status_var.set(f"Erro ao gerar relatório: {e}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao gerar o relatório:\n\n{e}"
            )
