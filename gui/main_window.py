import os
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
from datetime import datetime
import queue
import subprocess # Importar subprocess para TimeoutExpired

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
        # Define um tamanho inicial, mas permite maximizar
        self.root.geometry("1024x768") 
        self.root.minsize(800, 600)

        # Maximiza a janela ao iniciar (Linux)
        try:
            # Tenta usar o atributo específico para alguns WMs Linux
            self.root.attributes("-zoomed", True)
        except tk.TclError:
            # Fallback para o método padrão (pode não funcionar em todos WMs)
            try:
                self.root.state("zoomed")
            except tk.TclError:
                print("Aviso: Não foi possível maximizar a janela automaticamente.")
        
        # Flag para controlar se a coleta de hardware já foi iniciada
        self.hardware_collection_started = False
        self.hardware_collection_running = False
        
        # Flag para controlar se o cadastro foi concluído
        self.registration_completed = False
        
        # Flag para controlar se os testes estão em execução
        self.tests_running = False
        
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
        self.report_save_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.report_format_var = tk.StringVar(value="TXT")
        
        # Configura o estilo
        self._setup_style()
        
        # Cria a interface
        self._create_widgets()
        
        # Inicia a coleta de hardware em segundo plano
        self.root.after(100, self._start_hardware_collection)
    
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
            # Para sistemas não-Windows, verifica se é root
            if os.geteuid() != 0:
                messagebox.showwarning(
                    "Privilégios Insuficientes",
                    "Esta aplicação deve ser executada com privilégios de root (sudo) para "
                    "acessar todas as informações de hardware.\n\n"
                    "Algumas funcionalidades podem não estar disponíveis."
                )
    
    def _setup_style(self):
        """Configura o estilo da interface."""
        style = ttk.Style()
        
        # Configura o estilo para botões de sucesso
        style.configure("Success.TButton", background="green", foreground="white")
        
        # Configura o estilo para botões de ação
        style.configure("Action.TButton", font=("Arial", 10, "bold"))
        
        # Configura o estilo para labels de título
        style.configure("Title.TLabel", font=("Arial", 12, "bold"))
        
        # Configura o estilo para labels de subtítulo
        style.configure("Subtitle.TLabel", font=("Arial", 10, "bold"))
        
        # Estilos para abas
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", padding=[10, 5], font=("Arial", 10))
        
        # Estilos para frames
        style.configure("Card.TFrame", relief="raised", borderwidth=1)
        
        # Estilos para labels
        style.configure("Info.TLabel", font=("Arial", 10))
        style.configure("Value.TLabel", font=("Arial", 10, "bold"))
        
        # Estilos para botões
        style.configure("Primary.TButton", font=("Arial", 10, "bold"))
        style.configure("Secondary.TButton", font=("Arial", 10))
        style.configure("Warning.TButton", font=("Arial", 10, "bold"), background="#FFC107")
    
    def _create_widgets(self):
        """Cria os widgets da interface."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame superior para identificação
        top_frame = ttk.LabelFrame(main_frame, text="Identificação", padding=10)
        # Não expande verticalmente, fica fixo no topo
        top_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10)) 
        
        # Grid para campos de identificação
        ttk.Label(top_frame, text="Nome do Técnico:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.technician_entry = ttk.Entry(top_frame, textvariable=self.technician_name, width=30)
        self.technician_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(top_frame, text="ID da Bancada:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.workbench_entry = ttk.Entry(top_frame, textvariable=self.workbench_id, width=30)
        self.workbench_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Botão para concluir cadastro
        self.complete_registration_button = ttk.Button(
            top_frame,
            text="Concluir Cadastro",
            command=self._complete_registration,
            style="Action.TButton"
        )
        self.complete_registration_button.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky=tk.E)

        # Frame inferior para botões de ação (PACK ANTES DO NOTEBOOK)
        bottom_frame = ttk.Frame(main_frame)
        # Ancorado na parte inferior, não expande verticalmente
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0)) 
        
        # Notebook para abas (PACK DEPOIS DO TOP E BOTTOM FRAMES)
        self.notebook = ttk.Notebook(main_frame)
        # Expande para preencher o espaço restante entre top e bottom
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10) 
        
        # Aba de informações de hardware
        self.hardware_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.hardware_tab, text="Informações de Hardware")
        
        # Aba de testes
        self.tests_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tests_tab, text="Testes")
        
        # Aba de relatório
        self.report_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.report_tab, text="Relatório")
        
        # Configura a aba de informações de hardware
        self._setup_hardware_tab()
        
        # Configura a aba de testes
        self._setup_tests_tab()
        
        # Configura a aba de relatório
        self._setup_report_tab()
        
        # Adiciona os botões ao frame inferior
        # Botão para atualizar informações de hardware (Esquerda)
        self.refresh_button = ttk.Button(
            bottom_frame,
            text="Atualizar Informações",
            command=self._refresh_hardware_info,
            style="Action.TButton"
        )
        self.refresh_button.pack(side=tk.LEFT, padx=5)
        
        # Botão para executar testes selecionados (Centro)
        self.run_tests_button = ttk.Button(
            bottom_frame,
            text="Executar Testes Selecionados",
            command=self._run_selected_tests,
            state=tk.DISABLED,
            style="Action.TButton"
        )
        self.run_tests_button.pack(side=tk.LEFT, padx=5)
        
        # Botão para teste completo (Centro)
        self.run_all_button = ttk.Button(
            bottom_frame,
            text="Teste Completo",
            command=self._run_all_tests,
            state=tk.DISABLED,
            style="Primary.TButton"
        )
        self.run_all_button.pack(side=tk.LEFT, padx=5)

        # Botão para gerar relatório (Direita)
        self.generate_report_button = ttk.Button(
            bottom_frame,
            text="Gerar Relatório",
            command=self._generate_report,
            style="Action.TButton"
        )
        self.generate_report_button.pack(side=tk.RIGHT, padx=5)
    
    def _setup_hardware_tab(self):
        """Configura a aba de informações de hardware."""
        # Criar um canvas com scrollbar para a área de informações de hardware
        self.hardware_canvas = tk.Canvas(self.hardware_tab)
        self.hardware_scrollbar = ttk.Scrollbar(self.hardware_tab, orient="vertical", command=self.hardware_canvas.yview)
        
        # Configurar o canvas para usar a scrollbar
        self.hardware_canvas.configure(yscrollcommand=self.hardware_scrollbar.set)
        
        # Posicionar o canvas e a scrollbar
        self.hardware_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.hardware_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Criar um frame dentro do canvas para conter todos os widgets
        self.hardware_frame = ttk.Frame(self.hardware_canvas)
        self.hardware_canvas_window = self.hardware_canvas.create_window((0, 0), window=self.hardware_frame, anchor="nw")
        
        # Configurar o canvas para ajustar o tamanho do frame interno
        self.hardware_frame.bind("<Configure>", self._on_hardware_frame_configure)
        self.hardware_canvas.bind("<Configure>", self._on_hardware_canvas_configure)
        
        # Habilitar rolagem com a roda do mouse
        self.hardware_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Frame para informações da placa-mãe
        motherboard_frame = ttk.LabelFrame(self.hardware_frame, text="Placa-mãe", padding=10)
        motherboard_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(motherboard_frame, text="Fabricante:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.motherboard_manufacturer_var = tk.StringVar(value="Carregando...")
        ttk.Label(motherboard_frame, textvariable=self.motherboard_manufacturer_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(motherboard_frame, text="Modelo:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.motherboard_model_var = tk.StringVar(value="Carregando...")
        ttk.Label(motherboard_frame, textvariable=self.motherboard_model_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(motherboard_frame, text="Número de Série:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.motherboard_serial_var = tk.StringVar(value="Carregando...")
        ttk.Label(motherboard_frame, textvariable=self.motherboard_serial_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do processador
        cpu_frame = ttk.LabelFrame(self.hardware_frame, text="Processador", padding=10)
        cpu_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(cpu_frame, text="Marca:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.cpu_brand_var = tk.StringVar(value="Carregando...")
        ttk.Label(cpu_frame, textvariable=self.cpu_brand_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(cpu_frame, text="Modelo:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.cpu_model_var = tk.StringVar(value="Carregando...")
        ttk.Label(cpu_frame, textvariable=self.cpu_model_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações da memória RAM
        ram_frame = ttk.LabelFrame(self.hardware_frame, text="Memória RAM", padding=10)
        ram_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(ram_frame, text="Total:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.ram_total_var = tk.StringVar(value="Carregando...")
        ttk.Label(ram_frame, textvariable=self.ram_total_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(ram_frame, text="Slots Utilizados:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.ram_slots_var = tk.StringVar(value="Carregando...")
        ttk.Label(ram_frame, textvariable=self.ram_slots_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para módulos de memória
        self.ram_modules_frame = ttk.Frame(ram_frame)
        self.ram_modules_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Frame para informações dos discos
        disks_frame = ttk.LabelFrame(self.hardware_frame, text="Discos", padding=10)
        disks_frame.pack(fill=tk.X, pady=5)
        
        # Treeview para discos
        self.disk_tree = ttk.Treeview(disks_frame, columns=("model", "type", "size", "free"), show="headings", height=4)
        self.disk_tree.heading("model", text="Modelo")
        self.disk_tree.heading("type", text="Tipo")
        self.disk_tree.heading("size", text="Tamanho")
        self.disk_tree.heading("free", text="Livre")
        
        self.disk_tree.column("model", width=150)
        self.disk_tree.column("type", width=80)
        self.disk_tree.column("size", width=80)
        self.disk_tree.column("free", width=80)
        
        self.disk_tree.pack(fill=tk.BOTH, expand=True)
        
        # Frame para informações da placa de vídeo
        gpu_frame = ttk.LabelFrame(self.hardware_frame, text="Placa de Vídeo", padding=10)
        gpu_frame.pack(fill=tk.X, pady=5)
        
        # Treeview para GPUs
        self.gpu_tree = ttk.Treeview(gpu_frame, columns=("name"), show="headings", height=2)
        self.gpu_tree.heading("name", text="Modelo")
        self.gpu_tree.column("name", width=300)
        self.gpu_tree.pack(fill=tk.BOTH, expand=True)
        
        # Frame para informações do display
        display_frame = ttk.LabelFrame(self.hardware_frame, text="Display", padding=10)
        display_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(display_frame, text="Resolução:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.display_resolution_var = tk.StringVar(value="Carregando...")
        ttk.Label(display_frame, textvariable=self.display_resolution_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do TPM
        tpm_frame = ttk.LabelFrame(self.hardware_frame, text="TPM (Trusted Platform Module)", padding=10)
        tpm_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(tpm_frame, text="Versão:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.tpm_version_var = tk.StringVar(value="Carregando...")
        ttk.Label(tpm_frame, textvariable=self.tpm_version_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(tpm_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.tpm_status_var = tk.StringVar(value="Carregando...")
        ttk.Label(tpm_frame, textvariable=self.tpm_status_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(tpm_frame, text="Fabricante:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.tpm_manufacturer_var = tk.StringVar(value="Carregando...")
        ttk.Label(tpm_frame, textvariable=self.tpm_manufacturer_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do Bluetooth
        bluetooth_frame = ttk.LabelFrame(self.hardware_frame, text="Bluetooth", padding=10)
        bluetooth_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(bluetooth_frame, text="Dispositivo:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.bluetooth_device_var = tk.StringVar(value="Carregando...")
        ttk.Label(bluetooth_frame, textvariable=self.bluetooth_device_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(bluetooth_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.bluetooth_status_var = tk.StringVar(value="Carregando...")
        ttk.Label(bluetooth_frame, textvariable=self.bluetooth_status_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Frame para informações do Wi-Fi
        wifi_frame = ttk.LabelFrame(self.hardware_frame, text="Wi-Fi", padding=10)
        wifi_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(wifi_frame, text="Adaptador:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.wifi_adapter_var = tk.StringVar(value="Carregando...")
        ttk.Label(wifi_frame, textvariable=self.wifi_adapter_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(wifi_frame, text="Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.wifi_status_var = tk.StringVar(value="Carregando...")
        ttk.Label(wifi_frame, textvariable=self.wifi_status_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(wifi_frame, text="Rede Conectada:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.wifi_ssid_var = tk.StringVar(value="Carregando...")
        ttk.Label(wifi_frame, textvariable=self.wifi_ssid_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
    
    def _on_hardware_frame_configure(self, event):
        """Ajusta o tamanho do canvas quando o frame interno muda de tamanho."""
        self.hardware_canvas.configure(scrollregion=self.hardware_canvas.bbox("all"))
    
    def _on_hardware_canvas_configure(self, event):
        """Ajusta a largura do frame interno quando o canvas muda de tamanho."""
        self.hardware_canvas.itemconfig(self.hardware_canvas_window, width=event.width)
    
    def _on_mousewheel(self, event):
        """Permite rolagem com a roda do mouse."""
        # Verifica se o mouse está sobre o canvas de hardware
        # Apenas rola se o cursor estiver sobre o canvas correto
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is self.hardware_canvas or widget.master is self.hardware_canvas:
            # Calcula a direção da rolagem (diferente entre Linux e outros)
            if sys.platform == "linux":
                if event.num == 4:
                    delta = -1
                elif event.num == 5:
                    delta = 1
                else:
                    delta = 0
            else: # Windows, macOS
                delta = -1 * int(event.delta / 120)
            
            self.hardware_canvas.yview_scroll(delta, "units")

    
    def _setup_tests_tab(self):
        """Configura a aba de testes."""
        # Frame para seleção de testes
        tests_selection_frame = ttk.LabelFrame(self.tests_tab, text="Selecione os Testes", padding=10)
        tests_selection_frame.pack(fill=tk.X, pady=5)
        
        # Checkboxes para seleção de testes
        ttk.Checkbutton(
            tests_selection_frame,
            text="Teste de Teclado",
            variable=self.selected_tests["keyboard"]
        ).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            tests_selection_frame,
            text="Teste de USB",
            variable=self.selected_tests["usb"]
        ).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            tests_selection_frame,
            text="Teste de Webcam",
            variable=self.selected_tests["webcam"]
        ).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Checkbutton(
            tests_selection_frame,
            text="Teste de Áudio",
            variable=self.selected_tests["audio"]
        ).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Descrição dos testes
        description_frame = ttk.LabelFrame(self.tests_tab, text="Descrição dos Testes", padding=10)
        description_frame.pack(fill=tk.X, pady=5)
        
        descriptions = {
            "keyboard": "Teste de Teclado: Verifica o funcionamento de todas as teclas do teclado.",
            "usb": "Teste de USB: Verifica a conexão e velocidade de dispositivos USB.",
            "webcam": "Teste de Webcam: Verifica o funcionamento da webcam.",
            "audio": "Teste de Áudio: Verifica o funcionamento do microfone e alto-falantes."
        }
        
        for i, (test_id, description) in enumerate(descriptions.items()):
            ttk.Label(description_frame, text=description).pack(anchor=tk.W, padx=5, pady=2)
        
        # Frame para resultados dos testes
        self.tests_results_frame = ttk.LabelFrame(self.tests_tab, text="Resultados dos Testes", padding=10)
        self.tests_results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Texto para exibir os resultados
        self.tests_results_text = tk.Text(self.tests_results_frame, wrap=tk.WORD, height=20)
        self.tests_results_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Scrollbar para o texto
        scrollbar = ttk.Scrollbar(self.tests_results_frame, orient=tk.VERTICAL, command=self.tests_results_text.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        self.tests_results_text.config(yscrollcommand=scrollbar.set)
        
        # Configura o texto como somente leitura
        self.tests_results_text.config(state=tk.DISABLED)
    
    def _setup_report_tab(self):
        """Configura a aba de relatório."""
        # Frame para opções de relatório
        report_options_frame = ttk.LabelFrame(self.report_tab, text="Opções de Relatório", padding=10)
        report_options_frame.pack(fill=tk.X, pady=5)
        
        # Opções de relatório
        ttk.Label(report_options_frame, text="Formato:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Combobox para seleção de formato
        self.report_format_combo = ttk.Combobox(
            report_options_frame,
            textvariable=self.report_format_var,
            values=["TXT", "LOG"],
            state="readonly",
            width=10
        )
        self.report_format_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(report_options_frame, text="Salvar em:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.report_path_entry = ttk.Entry(report_options_frame, textvariable=self.report_save_path, width=50)
        self.report_path_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.browse_button = ttk.Button(
            report_options_frame,
            text="Procurar",
            command=self._browse_save_path
        )
        self.browse_button.grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Frame para visualização do relatório
        report_view_frame = ttk.LabelFrame(self.report_tab, text="Visualização do Relatório", padding=10)
        report_view_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Texto para exibir o relatório
        self.report_text = tk.Text(report_view_frame, wrap=tk.WORD, height=20)
        self.report_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Scrollbar para o texto
        scrollbar = ttk.Scrollbar(report_view_frame, orient=tk.VERTICAL, command=self.report_text.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        self.report_text.config(yscrollcommand=scrollbar.set)
        
        # Configura o texto como somente leitura
        self.report_text.config(state=tk.DISABLED)
    
    def _complete_registration(self):
        """Conclui o cadastro do técnico e bancada."""
        # Verifica se os campos foram preenchidos
        if not self.technician_name.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "Por favor, informe o nome do técnico."
            )
            return
        
        if not self.workbench_id.get().strip():
            messagebox.showwarning(
                "Campo Obrigatório",
                "Por favor, informe o ID da bancada."
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
            
            # Desabilita os campos e o botão
            self.technician_entry.config(state=tk.DISABLED)
            self.workbench_entry.config(state=tk.DISABLED)
            self.complete_registration_button.config(state=tk.DISABLED)
            
            # Habilita os botões de executar testes
            self.run_tests_button.config(state=tk.NORMAL)
            self.run_all_button.config(state=tk.NORMAL)
            
            # Exibe mensagem de sucesso
            messagebox.showinfo(
                "Cadastro Concluído",
                "Cadastro concluído com sucesso!\n\n"
                "Agora você pode selecionar e executar os testes."
            )
    
    def _start_hardware_collection(self):
        """Inicia a coleta de informações de hardware em segundo plano."""
        if not self.hardware_collection_started and not self.hardware_collection_running:
            # Define o estado como iniciado e rodando
            self.hardware_collection_started = True
            self.hardware_collection_running = True
            
            # Atualiza a interface para indicar carregamento
            self._set_loading_state(True)
            
            # Inicia a thread de coleta
            thread = threading.Thread(target=self._collect_hardware_info)
            thread.daemon = True
            thread.start()

    def _set_loading_state(self, loading):
        """Define o estado de carregamento para todos os campos de hardware."""
        state = "Carregando..." if loading else "Não disponível"
        # Placa-mãe
        self.motherboard_manufacturer_var.set(state)
        self.motherboard_model_var.set(state)
        self.motherboard_serial_var.set(state)
        # CPU
        self.cpu_brand_var.set(state)
        self.cpu_model_var.set(state)
        # RAM
        self.ram_total_var.set(state)
        self.ram_slots_var.set(state)
        for widget in self.ram_modules_frame.winfo_children():
            widget.destroy()
        if loading:
             ttk.Label(self.ram_modules_frame, text=state).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        # Discos
        self.disk_tree.delete(*self.disk_tree.get_children())
        if loading:
            self.disk_tree.insert("", 0, values=(state, state, state, state))
        # GPU
        self.gpu_tree.delete(*self.gpu_tree.get_children())
        if loading:
            self.gpu_tree.insert("", 0, values=(state,))
        # Display
        self.display_resolution_var.set(state)
        # TPM
        self.tpm_version_var.set(state)
        self.tpm_status_var.set(state)
        self.tpm_manufacturer_var.set(state)
        # Bluetooth
        self.bluetooth_device_var.set(state)
        self.bluetooth_status_var.set(state)
        # Wi-Fi
        self.wifi_adapter_var.set(state)
        self.wifi_status_var.set(state)
        self.wifi_ssid_var.set(state)

    def _collect_hardware_info(self):
        """Coleta as informações de hardware."""
        all_info = {}
        error_occurred = False
        error_message = ""
        try:
            # Coleta as informações
            # Usar get_all_info para simplificar e garantir que todos sejam chamados
            all_info = self.hardware_info.get_all_info()
            
        except subprocess.TimeoutExpired as e:
            print(f"Erro de Timeout ao coletar informações de hardware: {e}")
            error_occurred = True
            error_message = f"Timeout durante a coleta: {e.cmd}"
        except Exception as e:
            print(f"Erro geral ao coletar informações de hardware: {e}")
            error_occurred = True
            error_message = f"Erro inesperado: {e}"
        finally:
            # Garante que a flag seja resetada e a UI atualizada
            self.hardware_collection_running = False
            # Atualiza a interface no thread principal, passando as informações coletadas (ou vazias)
            self.root.after(0, lambda: self._update_hardware_info(all_info, error_occurred, error_message))
    
    def _update_hardware_info(self, all_info, error_occurred=False, error_message=""):
        """Atualiza a interface com as informações de hardware coletadas."""
        # Se ocorreu um erro grave durante a coleta, exibe aviso e usa "Erro" nos campos
        default_value = "Erro na coleta" if error_occurred else "Não disponível"
        if error_occurred:
             messagebox.showerror("Erro na Coleta", f"Não foi possível coletar todas as informações de hardware.\nDetalhes: {error_message}")

        # Extrai informações individuais ou usa dicionário vazio se não existir
        motherboard_info = all_info.get("motherboard", {})
        cpu_info = all_info.get("cpu", {})
        ram_info = all_info.get("ram", {})
        disks_info = all_info.get("disks", [])
        gpu_info = all_info.get("gpu", [])
        display_info = all_info.get("display", {})
        tpm_info = all_info.get("tpm", {})
        bluetooth_info = all_info.get("bluetooth", {})
        wifi_info = all_info.get("wifi", {})

        # Atualiza as informações da placa-mãe
        self.motherboard_manufacturer_var.set(motherboard_info.get("manufacturer", default_value))
        self.motherboard_model_var.set(motherboard_info.get("model", default_value))
        self.motherboard_serial_var.set(motherboard_info.get("serial_number", default_value)) 
        
        # Atualiza as informações do processador
        self.cpu_brand_var.set(cpu_info.get("brand", default_value))
        self.cpu_model_var.set(cpu_info.get("model", default_value))
        
        # Atualiza as informações da memória RAM
        self.ram_total_var.set(ram_info.get("total", default_value))
        self.ram_slots_var.set(ram_info.get("slots_used", default_value))
        
        # Limpa o frame de módulos de memória
        for widget in self.ram_modules_frame.winfo_children():
            widget.destroy()
        
        # Adiciona os módulos de memória
        modules = ram_info.get("modules", [])
        if modules:
            for i, module in enumerate(modules):
                if isinstance(module, dict):
                    module_text = f"Slot {i+1}: {module.get('size', 'N/A')} - {module.get('manufacturer', 'N/A')}"
                else:
                    module_text = f"Slot {i+1}: {module}"
                ttk.Label(self.ram_modules_frame, text=module_text).pack(anchor=tk.W)
        elif not error_occurred:
             ttk.Label(self.ram_modules_frame, text=default_value).pack(anchor=tk.W)

        # Limpa a treeview de discos
        self.disk_tree.delete(*self.disk_tree.get_children())
        
        # Adiciona os discos
        if disks_info:
            for i, disk in enumerate(disks_info):
                if isinstance(disk, dict):
                    self.disk_tree.insert(
                        "",
                        i,
                        values=(
                            disk.get('model', 'Não disponível'),
                            disk.get('type', 'Não disponível'),
                            disk.get('size', 'Não disponível'),
                            disk.get('free_space', 'Não disponível')
                        )
                    )
                else:
                    self.disk_tree.insert("", i, values=(str(disk), "N/A", "N/A", "N/A"))
        elif not error_occurred:
             self.disk_tree.insert("", 0, values=(default_value, default_value, default_value, default_value))

        # Limpa a treeview de GPUs
        self.gpu_tree.delete(*self.gpu_tree.get_children())
        
        # Adiciona as GPUs
        if gpu_info:
            for i, gpu in enumerate(gpu_info):
                if isinstance(gpu, dict):
                    self.gpu_tree.insert("", i, values=(gpu.get('name', 'Não disponível'),))
                else:
                    self.gpu_tree.insert("", i, values=(str(gpu),))
        elif not error_occurred:
             self.gpu_tree.insert("", 0, values=(default_value,))

        # Atualiza as informações do display
        self.display_resolution_var.set(display_info.get("resolution", default_value))
        
        # Atualiza as informações do TPM
        self.tpm_version_var.set(tpm_info.get("version", default_value))
        self.tpm_status_var.set(tpm_info.get("status", default_value))
        self.tpm_manufacturer_var.set(tpm_info.get("manufacturer", default_value))
        
        # Atualiza as informações do Bluetooth
        self.bluetooth_device_var.set(bluetooth_info.get("device", default_value))
        self.bluetooth_status_var.set(bluetooth_info.get("status", default_value))
        
        # Atualiza as informações do Wi-Fi
        self.wifi_adapter_var.set(wifi_info.get("adapter", default_value))
        self.wifi_status_var.set(wifi_info.get("status", default_value))
        self.wifi_ssid_var.set(wifi_info.get("ssid", default_value))
    
    def _refresh_hardware_info(self):
        """Atualiza as informações de hardware."""
        if not self.hardware_collection_running:
            # Define o estado como rodando
            self.hardware_collection_running = True
            
            # Atualiza a interface para indicar carregamento
            self._set_loading_state(True)
            
            # Inicia a thread de coleta
            thread = threading.Thread(target=self._collect_hardware_info)
            thread.daemon = True
            thread.start()
            
        else:
            messagebox.showwarning("Coleta em Andamento", "A coleta de informações de hardware já está em andamento.")

    def _run_selected_tests(self):
        """Executa os testes selecionados."""
        # Verifica se o cadastro foi concluído
        if not self.registration_completed:
            messagebox.showwarning(
                "Cadastro Incompleto",
                "Por favor, conclua o cadastro antes de executar os testes."
            )
            return
        
        # Verifica se algum teste foi selecionado
        selected = [test_id for test_id, var in self.selected_tests.items() if var.get()]
        
        if not selected:
            messagebox.showwarning(
                "Nenhum Teste Selecionado",
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
        # Cria o teste de teclado passando a janela principal como parent
        keyboard_test = KeyboardTest()
        
        if keyboard_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    keyboard_test.execute()
                    
                    # Registra o resultado
                    self.root.after(0, lambda: self._handle_test_completion(
                        "keyboard",
                        keyboard_test.get_result(),
                        keyboard_test.get_formatted_result()
                    ))
        else:
            error_message = keyboard_test.result.get("error", "Erro desconhecido ao inicializar o teste de teclado.")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de teclado:\n\n{error_message}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de teclado:\n\n{keyboard_test.result["error"]}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)    
    def _run_usb_test(self):
        """Executa o teste de USB."""
        # Cria o teste de USB
        usb_test = USBTest()
        
        if usb_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    usb_test.execute()
                    
                    # Registra o resultado
                    self.root.after(0, lambda: self._handle_test_completion(
                        "usb",
                        usb_test.get_result(),
                        usb_test.get_formatted_result()
                    ))                except Exception as e:
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de USB:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            error_message = usb_test.result.get("error", "Erro desconhecido ao inicializar o teste de USB.")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de USB:\n\n{error_message}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)    
    def _run_webcam_test(self):
        """Executa o teste de webcam."""
        webcam_test = WebcamTest()
        
        if webcam_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    webcam_test.execute()
                    
                    # Registra o resultado
                    self.root.after(0, lambda: self._handle_test_completion(
                        "Webcam",
                        webcam_test.get_result(),
                        webcam_test.get_formatted_result()
                    ))                except Exception as e:
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de webcam:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            error_message = webcam_test.result.get("error", "Erro desconhecido ao inicializar o teste de webcam.")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de webcam:\n\n{error_message}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)

    def _run_audio_test(self):
        """Executa o teste de áudio."""
        audio_test = AudioTest()
        
        if audio_test.initialize():
            # Executa o teste em uma thread separada
            def run_test():
                try:
                    audio_test.execute()
                    
                    # Registra o resultado
                    self.root.after(0, lambda: self._handle_test_completion(
                        "Áudio",
                        audio_test.get_result(),
                        audio_test.get_formatted_result()
                    ))                except Exception as e:
                    messagebox.showerror(
                        "Erro",
                        f"Ocorreu um erro ao executar o teste de áudio:\n\n{e}"
                    )
                    
                    # Executa o próximo teste
                    self.root.after(100, self._execute_next_test)
            
            # Inicia a thread
            threading.Thread(target=run_test, daemon=True).start()
        else:
            error_message = audio_test.result.get("error", "Erro desconhecido ao inicializar o teste de áudio.")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao inicializar o teste de áudio:\n\n{error_message}"
            )
            
            # Executa o próximo teste
            self.root.after(100, self._execute_next_test)


    def _generate_report(self):
        """Gera o relatório."""
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

        try:
            # Define as informações de identificação
            identification_info = {
                "date_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "technician_name": self.technician_name.get(),
                "workbench_id": self.workbench_id.get()
            }
            self.report_generator.set_identification(identification_info)

            # Define as informações de hardware
            hardware_info = {
                'motherboard': {
                    'manufacturer': self.motherboard_manufacturer_var.get(),
                    'model': self.motherboard_model_var.get(),
                    'serial_number': self.motherboard_serial_var.get()
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
                    'device_name': self.bluetooth_device_var.get(),
                    'device_status': self.bluetooth_status_var.get()
                },
                'wifi': {
                    'adapter_name': self.wifi_adapter_var.get(),
                    'adapter_status': self.wifi_status_var.get(),
                    'connected_ssid': self.wifi_ssid_var.get()
                }
            }
            self.report_generator.set_hardware_info(hardware_info)

            # Define o diretório de saída
            self.report_generator.set_output_directory(self.report_save_path.get())

            # Gera o relatório
            report_path = self.report_generator.generate_report()

            # Exibe o relatório na aba de visualização
            with open(report_path, 'r', encoding='utf-8') as f:
                report_content = f.read()
            self.report_text.config(state=tk.NORMAL)
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(tk.END, report_content)
            self.report_text.config(state=tk.DISABLED)

            messagebox.showinfo(
                "Relatório Gerado",
                f"Relatório gerado com sucesso!\n\nSalvo em: {report_path}"
            )

        except Exception as e:
            messagebox.showerror("Erro ao Gerar Relatório", f"Ocorreu um erro ao gerar o relatório: {e}")

    def _browse_save_path(self):
        """Permite ao usuário selecionar o diretório para salvar o relatório."""
        from tkinter import filedialog
        selected_path = filedialog.askdirectory(
            parent=self.root,
            initialdir=self.report_save_path.get(),
            title="Selecionar Pasta para Salvar Relatório"
        )
        if selected_path:
            self.report_save_path.set(selected_path)



    def _handle_test_completion(self, test_name, result, formatted_result):
        """Callback para quando um teste é concluído e para atualizar a interface."""
        # Mapeamento para nomes amigáveis
        traducao = {
            "keyboard": "Teclado",
            "usb": "USB",
            "webcam": "Webcam",
            "audio": "Áudio"
        }
        nome_teste = traducao.get(test_name, test_name.capitalize())

        # Resultado amigável
        if result.get("success"):
            resultado_str = "Concluído com Sucesso"
        else:
            resultado_str = "Falha"

        datahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        texto = (
            f"Teste: {nome_teste}\n"
            f"Resultado: {resultado_str}\n"
            f"Detalhes: {formatted_result}\n"
            f"Data/Hora: {datahora}\n\n"
        )

        # Exibe no widget de resultados
        self.tests_results_text.config(state=tk.NORMAL)
        self.tests_results_text.insert(tk.END, texto)
        self.tests_results_text.config(state=tk.DISABLED)

        # Salva para o relatório
        self.report_generator.add_test_result(test_name, result, formatted_result)
        
        # Executa o próximo teste
        self._execute_next_test()