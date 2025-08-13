"""
Módulo para teste de Áudio.
Implementa teste real de áudio com gravação e reprodução de som.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import tempfile
import wave
import numpy as np
import sounddevice as sd

class AudioTest:
    """Classe para teste de áudio."""
    
    def __init__(self):
        """Inicializa o teste de áudio."""
        self.window = None
        self.is_running = False
        self.is_recording = False
        self.is_playing = False
        self.is_completed = False
        self.result = {
            'success': False,
            'message': '',
            'details': {},
            'error': None
        }
        self.temp_dir = None
        self.audio_file = None
        self.recorded_frames = []
        self.sample_rate = 44100
        self.channels = 2
        self.recording_duration = 10  # segundos
    
    def initialize(self):
        """Inicializa o teste de áudio."""
        try:
            # Cria um diretório temporário para salvar o áudio
            self.temp_dir = tempfile.mkdtemp()
            self.audio_file = os.path.join(self.temp_dir, "audio_test.wav")
            
            self.is_running = True
            return True
        except Exception as e:
            self.result['error'] = str(e)
            return False
    
    def execute(self):
        """Executa o teste de áudio."""
        try:
            # Cria a janela de teste
            self.window = tk.Toplevel()
            self.window.title("Teste de Áudio")
            self.window.geometry("600x400")
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
                text="Teste de Áudio",
                font=("Arial", 14, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            # Frame para teste de microfone
            mic_frame = ttk.LabelFrame(main_frame, text="Teste de Microfone", padding=10)
            mic_frame.pack(fill=tk.X, pady=10)
            
            # Botão para iniciar/parar gravação
            self.record_button = ttk.Button(
                mic_frame,
                text="Iniciar Gravação (10s)",
                command=self._toggle_recording
            )
            self.record_button.pack(pady=5)
            
            # Barra de progresso para gravação
            self.record_progress_var = tk.DoubleVar(value=0)
            self.record_progress = ttk.Progressbar(
                mic_frame,
                orient=tk.HORIZONTAL,
                length=100,
                mode='determinate',
                variable=self.record_progress_var
            )
            self.record_progress.pack(fill=tk.X, pady=5)
            
            # Status da gravação
            self.record_status_var = tk.StringVar(value="Clique em 'Iniciar Gravação' para começar o teste")
            record_status_label = ttk.Label(mic_frame, textvariable=self.record_status_var)
            record_status_label.pack(fill=tk.X, pady=5)
            
            # Frame para teste de reprodução
            play_frame = ttk.LabelFrame(main_frame, text="Teste de Reprodução", padding=10)
            play_frame.pack(fill=tk.X, pady=10)
            
            # Botão para reproduzir áudio
            self.play_button = ttk.Button(
                play_frame,
                text="Reproduzir Áudio",
                state=tk.DISABLED,
                command=self._play_audio
            )
            self.play_button.pack(pady=5)
            
            # Status da reprodução
            self.play_status_var = tk.StringVar(value="Grave um áudio primeiro")
            play_status_label = ttk.Label(play_frame, textvariable=self.play_status_var)
            play_status_label.pack(fill=tk.X, pady=5)
            
            # Frame para confirmação
            confirm_frame = ttk.LabelFrame(main_frame, text="Confirmação", padding=10)
            confirm_frame.pack(fill=tk.X, pady=10)
            
            # Pergunta de confirmação
            confirm_label = ttk.Label(
                confirm_frame,
                text="O teste de áudio funcionou corretamente?"
            )
            confirm_label.pack(pady=5)
            
            # Frame para botões de confirmação
            buttons_frame = ttk.Frame(confirm_frame)
            buttons_frame.pack(fill=tk.X, pady=5)
            
            # Botão Sim
            self.yes_button = ttk.Button(
                buttons_frame,
                text="Sim",
                state=tk.DISABLED,
                command=lambda: self._confirm_audio(True)
            )
            self.yes_button.pack(side=tk.LEFT, padx=5)
            
            # Botão Não
            self.no_button = ttk.Button(
                buttons_frame,
                text="Não",
                state=tk.DISABLED,
                command=lambda: self._confirm_audio(False)
            )
            self.no_button.pack(side=tk.RIGHT, padx=5)
            
            # Frame para botões de ação
            action_frame = ttk.Frame(main_frame)
            action_frame.pack(fill=tk.X, pady=(10, 0))
            
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
    
    def _toggle_recording(self):
        """Inicia ou para a gravação de áudio."""
        if not self.is_recording:
            # Inicia a gravação
            self._start_recording()
        else:
            # Para a gravação
            self._stop_recording()
    
    def _start_recording(self):
        """Inicia a gravação de áudio."""
        # Limpa os frames anteriores
        self.recorded_frames = []
        
        # Atualiza o status
        self.record_status_var.set("Iniciando gravação...")
        
        # Atualiza o botão
        self.record_button.config(text="Parar Gravação")
        
        # Desabilita o botão de reprodução
        self.play_button.config(state=tk.DISABLED)
        
        # Marca que está gravando
        self.is_recording = True
        
        # Inicia a gravação em uma thread separada
        threading.Thread(target=self._record_audio, daemon=True).start()
    
    def _record_audio(self):
        """Grava áudio do microfone."""
        try:
            # Configura o callback de gravação
            def callback(indata, frames, time, status):
                if status:
                    print(f"Status: {status}")
                self.recorded_frames.append(indata.copy())
            
            # Inicia a gravação
            with sd.InputStream(samplerate=self.sample_rate, channels=self.channels, callback=callback):
                # Atualiza o status
                self.window.after(0, lambda: self.record_status_var.set("Gravando..."))
                
                # Atualiza a barra de progresso
                for i in range(self.recording_duration * 10):
                    if not self.is_recording:
                        break
                    
                    progress = (i + 1) / (self.recording_duration * 10) * 100
                    self.window.after(0, lambda p=progress: self.record_progress_var.set(p))
                    
                    time.sleep(0.1)
                
                # Se ainda estiver gravando, para automaticamente
                if self.is_recording:
                    self.window.after(0, self._stop_recording)
        except Exception as e:
            # Atualiza o status
            self.window.after(0, lambda: self.record_status_var.set(f"Erro: {e}"))
            
            # Armazena o erro
            self.result['error'] = str(e)
            
            # Marca que não está mais gravando
            self.is_recording = False
            
            # Atualiza o botão
            self.window.after(0, lambda: self.record_button.config(text="Iniciar Gravação (10s)"))
    
    def _stop_recording(self):
        """Para a gravação de áudio."""
        # Marca que não está mais gravando
        self.is_recording = False
        
        # Atualiza o botão
        self.record_button.config(text="Iniciar Gravação (10s)")
        
        # Atualiza o status
        self.record_status_var.set("Gravação concluída")
        
        # Salva o áudio gravado
        if self.recorded_frames:
            try:
                # Converte os frames para um array numpy
                recorded_data = np.concatenate(self.recorded_frames, axis=0)
                
                # Salva o arquivo WAV
                with wave.open(self.audio_file, 'wb') as wf:
                    wf.setnchannels(self.channels)
                    wf.setsampwidth(2)  # 16 bits
                    wf.setframerate(self.sample_rate)
                    wf.writeframes((recorded_data * 32767).astype(np.int16).tobytes())
                
                # Habilita o botão de reprodução
                self.play_button.config(state=tk.NORMAL)
                
                # Atualiza o status de reprodução
                self.play_status_var.set("Clique em 'Reproduzir Áudio' para ouvir a gravação")
            except Exception as e:
                # Atualiza o status
                self.record_status_var.set(f"Erro ao salvar áudio: {e}")
                
                # Armazena o erro
                self.result['error'] = str(e)
        else:
            # Atualiza o status
            self.record_status_var.set("Nenhum áudio gravado")
    
    def _play_audio(self):
        """Reproduz o áudio gravado."""
        if not os.path.exists(self.audio_file):
            self.play_status_var.set("Arquivo de áudio não encontrado")
            return
        
        if self.is_playing:
            self.play_status_var.set("Já está reproduzindo áudio")
            return
        
        # Marca que está reproduzindo
        self.is_playing = True
        
        # Atualiza o status
        self.play_status_var.set("Reproduzindo áudio...")
        
        # Desabilita o botão de reprodução
        self.play_button.config(state=tk.DISABLED)
        
        # Inicia a reprodução em uma thread separada
        threading.Thread(target=self._play_audio_thread, daemon=True).start()
    
    def _play_audio_thread(self):
        """Thread para reprodução de áudio."""
        try:
            # Abre o arquivo WAV
            with wave.open(self.audio_file, 'rb') as wf:
                # Configura o stream de áudio
                sd.play(
                    sd.read_wav(self.audio_file)[0],
                    samplerate=wf.getframerate()
                )
                sd.wait()
            
            # Atualiza o status
            self.window.after(0, lambda: self.play_status_var.set("Reprodução concluída"))
            
            # Habilita os botões de confirmação
            self.window.after(0, lambda: self.yes_button.config(state=tk.NORMAL))
            self.window.after(0, lambda: self.no_button.config(state=tk.NORMAL))
        except Exception as e:
            # Atualiza o status
            self.window.after(0, lambda: self.play_status_var.set(f"Erro: {e}"))
            
            # Armazena o erro
            self.result['error'] = str(e)
        finally:
            # Marca que não está mais reproduzindo
            self.is_playing = False
            
            # Habilita o botão de reprodução
            self.window.after(0, lambda: self.play_button.config(state=tk.NORMAL))
    
    def _confirm_audio(self, success):
        """Confirma o funcionamento do áudio."""
        # Armazena o resultado
        self.result['success'] = success
        
        if success:
            self.result['message'] = "Áudio funcionou corretamente"
            self.play_status_var.set("Áudio funcionou corretamente. Clique em 'Concluir Teste' para finalizar.")
        else:
            self.result['message'] = "Áudio não funcionou corretamente"
            self.play_status_var.set("Áudio não funcionou corretamente. Clique em 'Concluir Teste' para finalizar.")
        
        # Habilita o botão de concluir
        self.complete_button.config(state=tk.NORMAL)
    
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
            "Tem certeza que deseja pular o teste de áudio?\n\n"
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
            "Tem certeza que deseja fechar o teste de áudio?\n\n"
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
            return f"Teste de Áudio: SUCESSO\n" \
                   f"Mensagem: {self.result['message']}"
        else:
            if self.result['error']:
                return f"Teste de Áudio: FALHA\n" \
                       f"Erro: {self.result['error']}"
            else:
                return f"Teste de Áudio: FALHA\n" \
                       f"Motivo: {self.result['message']}"
