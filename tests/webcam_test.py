"""
Módulo para teste de Webcam.
Implementa teste real de webcam com captura de imagem e confirmação do usuário.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import cv2
from PIL import Image, ImageTk
import tempfile

class WebcamTest:
    """Classe para teste de webcam."""
    
    def __init__(self):
        """Inicializa o teste de webcam."""
        self.window = None
        self.cap = None
        self.is_running = False
        self.is_capturing = False
        self.is_completed = False
        self.result = {
            'success': False,
            'message': '',
            'details': {},
            'error': None
        }
        self.temp_dir = None
        self.photo_path = None
    
    def initialize(self):
        """Inicializa o teste de webcam."""
        try:
            # Cria um diretório temporário para salvar a foto
            self.temp_dir = tempfile.mkdtemp()
            self.photo_path = os.path.join(self.temp_dir, "webcam_test.jpg")
            
            self.is_running = True
            return True
        except Exception as e:
            self.result['error'] = str(e)
            return False
    
    def execute(self):
        """Executa o teste de webcam."""
        try:
            # Cria a janela de teste
            self.window = tk.Toplevel()
            self.window.title("Teste de Webcam")
            self.window.geometry("800x600")
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
                text="Teste de Webcam",
                font=("Arial", 14, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            # Frame para exibição da webcam
            self.video_frame = ttk.Frame(main_frame)
            self.video_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # Label para exibição do vídeo
            self.video_label = ttk.Label(self.video_frame)
            self.video_label.pack(fill=tk.BOTH, expand=True)
            
            # Status
            self.status_var = tk.StringVar(value="Clique em 'Iniciar Webcam' para começar o teste")
            status_label = ttk.Label(main_frame, textvariable=self.status_var)
            status_label.pack(fill=tk.X, pady=5)
            
            # Frame para botões de ação
            action_frame = ttk.Frame(main_frame)
            action_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Botão para iniciar a webcam
            self.start_button = ttk.Button(
                action_frame,
                text="Iniciar Webcam",
                command=self._start_webcam
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
            # Garante que a webcam seja liberada
            self._release_webcam()
            
            # Limpa os arquivos temporários
            self._cleanup_temp_files()
    
    def _start_webcam(self):
        """Inicia a captura da webcam."""
        # Desabilita o botão de iniciar
        self.start_button.config(state=tk.DISABLED)
        
        # Inicia a captura em uma thread separada
        threading.Thread(target=self._capture_webcam, daemon=True).start()
    
    def _capture_webcam(self):
        """Captura o vídeo da webcam."""
        try:
            # Inicializa a captura
            self.cap = cv2.VideoCapture(0)
            
            if not self.cap.isOpened():
                raise Exception("Não foi possível acessar a webcam")
            
            # Atualiza o status
            self.window.after(0, lambda: self.status_var.set("Webcam iniciada. Exibindo imagem por 5 segundos..."))
            
            # Marca que está capturando
            self.is_capturing = True
            
            # Contador para exibição por 5 segundos
            start_time = time.time()
            
            # Loop de captura
            while self.is_capturing and time.time() - start_time < 5:
                # Captura um frame
                ret, frame = self.cap.read()
                
                if not ret:
                    raise Exception("Erro ao capturar frame da webcam")
                
                # Converte o frame para RGB
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                
                # Redimensiona o frame para caber na janela
                height, width, _ = frame_rgb.shape
                max_width = self.video_frame.winfo_width()
                max_height = self.video_frame.winfo_height()
                
                if max_width > 0 and max_height > 0:
                    # Calcula a proporção para manter a relação de aspecto
                    ratio = min(max_width / width, max_height / height)
                    new_width = int(width * ratio)
                    new_height = int(height * ratio)
                    
                    # Redimensiona o frame
                    frame_resized = cv2.resize(frame_rgb, (new_width, new_height))
                    
                    # Converte o frame para imagem do Tkinter
                    img = Image.fromarray(frame_resized)
                    img_tk = ImageTk.PhotoImage(image=img)
                    
                    # Atualiza a label
                    self.window.after(0, lambda i=img_tk: self._update_video_label(i))
                
                # Aguarda um pouco
                time.sleep(0.03)
            
            # Salva o último frame como foto
            if ret:
                cv2.imwrite(self.photo_path, frame)
            
            # Libera a webcam
            self._release_webcam()
            
            # Pergunta ao usuário se a webcam funcionou corretamente
            self.window.after(0, self._ask_webcam_confirmation)
        except Exception as e:
            # Atualiza o status
            self.window.after(0, lambda: self.status_var.set(f"Erro: {e}"))
            
            # Armazena o erro
            self.result['error'] = str(e)
            
            # Habilita o botão de iniciar
            self.window.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            
            # Libera a webcam
            self._release_webcam()
    
    def _update_video_label(self, img_tk):
        """Atualiza a label de vídeo com a imagem capturada."""
        self.video_label.configure(image=img_tk)
        self.video_label.image = img_tk
    
    def _ask_webcam_confirmation(self):
        """Pergunta ao usuário se a webcam funcionou corretamente."""
        # Atualiza o status
        self.status_var.set("A webcam funcionou corretamente?")
        
        # Cria um frame para os botões de confirmação
        confirm_frame = ttk.Frame(self.video_frame)
        confirm_frame.pack(fill=tk.X, pady=10)
        
        # Botão Sim
        yes_button = ttk.Button(
            confirm_frame,
            text="Sim",
            command=lambda: self._confirm_webcam(True)
        )
        yes_button.pack(side=tk.LEFT, padx=5)
        
        # Botão Não
        no_button = ttk.Button(
            confirm_frame,
            text="Não",
            command=lambda: self._confirm_webcam(False)
        )
        no_button.pack(side=tk.RIGHT, padx=5)
    
    def _confirm_webcam(self, success):
        """Confirma o funcionamento da webcam."""
        # Armazena o resultado
        self.result['success'] = success
        
        if success:
            self.result['message'] = "Webcam funcionou corretamente"
            self.status_var.set("Webcam funcionou corretamente. Clique em 'Concluir Teste' para finalizar.")
        else:
            self.result['message'] = "Webcam não funcionou corretamente"
            self.status_var.set("Webcam não funcionou corretamente. Clique em 'Concluir Teste' para finalizar.")
        
        # Habilita o botão de concluir
        self.complete_button.config(state=tk.NORMAL)
    
    def _release_webcam(self):
        """Libera a webcam."""
        self.is_capturing = False
        
        if self.cap:
            self.cap.release()
            self.cap = None
    
    def _cleanup_temp_files(self):
        """Limpa os arquivos temporários."""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                import shutil
                shutil.rmtree(self.temp_dir, ignore_errors=True)
        except:
            pass
    
    def _skip_test(self):
        """Pula o teste."""
        if messagebox.askyesno(
            "Pular Teste",
            "Tem certeza que deseja pular o teste de webcam?\n\n"
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
            "Tem certeza que deseja fechar o teste de webcam?\n\n"
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
            return f"Teste de Webcam: SUCESSO\n" \
                   f"Mensagem: {self.result['message']}"
        else:
            if self.result['error']:
                return f"Teste de Webcam: FALHA\n" \
                       f"Erro: {self.result['error']}"
            else:
                return f"Teste de Webcam: FALHA\n" \
                       f"Motivo: {self.result['message']}"
