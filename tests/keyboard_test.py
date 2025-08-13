"""
Módulo para teste de Teclado.
Implementa teste real de teclado com interface visual similar ao testarteclado.com.br.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from pynput import keyboard

class KeyboardTest:
    """Classe para teste de teclado."""
    
    def __init__(self):
        """Inicializa o teste de teclado."""
        self.window = None
        self.key_buttons = {}
        self.pressed_keys = set()
        self.listener = None
        self.is_running = False
        self.is_completed = False
        self.result = {
            'success': False,
            'message': '',
            'details': {},
            'error': None
        }
    
    def initialize(self):
        """Inicializa o teste de teclado."""
        try:
            self.is_running = True
            return True
        except Exception as e:
            self.result['error'] = str(e)
            return False
    
    def execute(self):
        """Executa o teste de teclado."""
        try:
            # Cria a janela de teste
            self.window = tk.Toplevel()
            self.window.title("Teste de Teclado")
            self.window.geometry("1000x400")
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
                text="Teste de Teclado - Pressione todas as teclas para validar",
                font=("Arial", 14, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            # Contador de teclas
            self.counter_var = tk.StringVar(value="0 de 0 teclas pressionadas")
            counter_label = ttk.Label(
                main_frame,
                textvariable=self.counter_var,
                font=("Arial", 12)
            )
            counter_label.pack(pady=(0, 10))
            
            # Frame para o teclado
            keyboard_frame = ttk.Frame(main_frame)
            keyboard_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # Cria o layout do teclado ABNT2
            self._create_keyboard_layout(keyboard_frame)
            
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
            
            # Inicia o listener de teclado
            self.listener = keyboard.Listener(
                on_press=self._on_key_press,
                on_release=self._on_key_release
            )
            self.listener.start()
            
            # Aguarda a conclusão do teste
            self.window.wait_window()
            
            # Retorna o resultado
            return self.is_completed
        except Exception as e:
            self.result['error'] = str(e)
            return False
        finally:
            # Garante que o listener seja encerrado
            if self.listener:
                self.listener.stop()
    
    def _create_keyboard_layout(self, parent):
        """Cria o layout do teclado ABNT2."""
        # Define as teclas do teclado ABNT2
        keyboard_layout = [
            # Primeira linha
            [('Esc', 1), ('F1', 1), ('F2', 1), ('F3', 1), ('F4', 1), ('F5', 1), ('F6', 1), ('F7', 1), ('F8', 1), ('F9', 1), ('F10', 1), ('F11', 1), ('F12', 1), ('PrtSc', 1), ('ScrLk', 1), ('Pause', 1)],
            # Segunda linha
            [('`', 1), ('1', 1), ('2', 1), ('3', 1), ('4', 1), ('5', 1), ('6', 1), ('7', 1), ('8', 1), ('9', 1), ('0', 1), ('-', 1), ('=', 1), ('Backspace', 2), ('Insert', 1), ('Home', 1), ('PgUp', 1)],
            # Terceira linha
            [('Tab', 1.5), ('Q', 1), ('W', 1), ('E', 1), ('R', 1), ('T', 1), ('Y', 1), ('U', 1), ('I', 1), ('O', 1), ('P', 1), ('[', 1), (']', 1), ('\\', 1.5), ('Delete', 1), ('End', 1), ('PgDn', 1)],
            # Quarta linha
            [('Caps Lock', 1.75), ('A', 1), ('S', 1), ('D', 1), ('F', 1), ('G', 1), ('H', 1), ('J', 1), ('K', 1), ('L', 1), ('Ç', 1), (';', 1), ('\'', 1), ('Enter', 2.25)],
            # Quinta linha
            [('Shift', 1.25), ('\\', 1), ('Z', 1), ('X', 1), ('C', 1), ('V', 1), ('B', 1), ('N', 1), ('M', 1), (',', 1), ('.', 1), ('/', 1), ('Shift', 2.75), ('↑', 1)],
            # Sexta linha
            [('Ctrl', 1.25), ('Win', 1.25), ('Alt', 1.25), ('Space', 6.25), ('Alt Gr', 1.25), ('Win', 1.25), ('Menu', 1.25), ('Ctrl', 1.25), ('←', 1), ('↓', 1), ('→', 1)]
        ]
        
        # Cria os botões para cada tecla
        for row_idx, row in enumerate(keyboard_layout):
            row_frame = ttk.Frame(parent)
            row_frame.pack(fill=tk.X, pady=2)
            
            for key, width in row:
                button = ttk.Button(
                    row_frame,
                    text=key,
                    width=int(width * 4)
                )
                button.pack(side=tk.LEFT, padx=1, pady=1)
                
                # Armazena o botão para referência futura
                self.key_buttons[key] = button
        
        # Atualiza o contador de teclas
        self.counter_var.set(f"0 de {len(self.key_buttons)} teclas pressionadas")
    
    def _on_key_press(self, key):
        """Manipulador de evento de tecla pressionada."""
        if not self.is_running or not self.window:
            return
        
        try:
            # Converte a tecla para string
            key_str = self._get_key_string(key)
            
            # Atualiza o conjunto de teclas pressionadas
            if key_str and key_str in self.key_buttons:
                self.pressed_keys.add(key_str)
                
                # Atualiza o estilo do botão
                self.window.after(0, lambda k=key_str: self._update_button_style(k, True))
                
                # Atualiza o contador
                self.window.after(0, self._update_counter)
                
                # Verifica se todas as teclas foram pressionadas
                if len(self.pressed_keys) == len(self.key_buttons):
                    self.window.after(0, lambda: self.complete_button.config(state=tk.NORMAL))
        except Exception as e:
            print(f"Erro ao processar tecla pressionada: {e}")
    
    def _on_key_release(self, key):
        """Manipulador de evento de tecla liberada."""
        # Não faz nada, apenas mantém o registro das teclas pressionadas
        pass
    
    def _get_key_string(self, key):
        """Converte uma tecla para string."""
        try:
            # Teclas especiais
            if hasattr(key, 'name'):
                if key.name == 'space':
                    return 'Space'
                elif key.name == 'shift':
                    return 'Shift'
                elif key.name == 'shift_r':
                    return 'Shift'
                elif key.name == 'ctrl':
                    return 'Ctrl'
                elif key.name == 'ctrl_r':
                    return 'Ctrl'
                elif key.name == 'alt':
                    return 'Alt'
                elif key.name == 'alt_gr':
                    return 'Alt Gr'
                elif key.name == 'menu':
                    return 'Menu'
                elif key.name == 'cmd':
                    return 'Win'
                elif key.name == 'cmd_r':
                    return 'Win'
                elif key.name == 'esc':
                    return 'Esc'
                elif key.name == 'tab':
                    return 'Tab'
                elif key.name == 'caps_lock':
                    return 'Caps Lock'
                elif key.name == 'enter':
                    return 'Enter'
                elif key.name == 'backspace':
                    return 'Backspace'
                elif key.name == 'delete':
                    return 'Delete'
                elif key.name == 'insert':
                    return 'Insert'
                elif key.name == 'home':
                    return 'Home'
                elif key.name == 'end':
                    return 'End'
                elif key.name == 'page_up':
                    return 'PgUp'
                elif key.name == 'page_down':
                    return 'PgDn'
                elif key.name == 'print_screen':
                    return 'PrtSc'
                elif key.name == 'scroll_lock':
                    return 'ScrLk'
                elif key.name == 'pause':
                    return 'Pause'
                elif key.name == 'up':
                    return '↑'
                elif key.name == 'down':
                    return '↓'
                elif key.name == 'left':
                    return '←'
                elif key.name == 'right':
                    return '→'
                elif key.name.startswith('f') and key.name[1:].isdigit():
                    return key.name.upper()
                else:
                    return None
            
            # Teclas normais
            char = key.char
            if char:
                if char == '`' or char == '~':
                    return '`'
                elif char == '1' or char == '!':
                    return '1'
                elif char == '2' or char == '@':
                    return '2'
                elif char == '3' or char == '#':
                    return '3'
                elif char == '4' or char == '$':
                    return '4'
                elif char == '5' or char == '%':
                    return '5'
                elif char == '6' or char == '¨':
                    return '6'
                elif char == '7' or char == '&':
                    return '7'
                elif char == '8' or char == '*':
                    return '8'
                elif char == '9' or char == '(':
                    return '9'
                elif char == '0' or char == ')':
                    return '0'
                elif char == '-' or char == '_':
                    return '-'
                elif char == '=' or char == '+':
                    return '='
                elif char == '[' or char == '{':
                    return '['
                elif char == ']' or char == '}':
                    return ']'
                elif char == '\\' or char == '|':
                    return '\\'
                elif char == ';' or char == ':':
                    return ';'
                elif char == '\'' or char == '"':
                    return '\''
                elif char == ',' or char == '<':
                    return ','
                elif char == '.' or char == '>':
                    return '.'
                elif char == '/' or char == '?':
                    return '/'
                elif char.upper() == 'Ç':
                    return 'Ç'
                else:
                    return char.upper()
        except:
            return None
    
    def _update_button_style(self, key, pressed):
        """Atualiza o estilo do botão."""
        if key in self.key_buttons:
            button = self.key_buttons[key]
            if pressed:
                button.configure(style='Success.TButton')
            else:
                button.configure(style='TButton')
    
    def _update_counter(self):
        """Atualiza o contador de teclas pressionadas."""
        self.counter_var.set(f"{len(self.pressed_keys)} de {len(self.key_buttons)} teclas pressionadas")
    
    def _skip_test(self):
        """Pula o teste."""
        if messagebox.askyesno(
            "Pular Teste",
            "Tem certeza que deseja pular o teste de teclado?\n\n"
            "O teste será marcado como não concluído."
        ):
            self.result['success'] = False
            self.result['message'] = "Teste pulado pelo usuário"
            self.is_completed = False
            self.window.destroy()
    
    def _complete_test(self):
        """Conclui o teste."""
        self.result['success'] = True
        self.result['message'] = "Teste concluído com sucesso"
        self.result['details'] = {
            'total_keys': len(self.key_buttons),
            'pressed_keys': len(self.pressed_keys)
        }
        self.is_completed = True
        self.window.destroy()
    
    def _on_close(self):
        """Manipulador de evento de fechamento da janela."""
        if messagebox.askyesno(
            "Fechar Teste",
            "Tem certeza que deseja fechar o teste de teclado?\n\n"
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
            return f"Teste de Teclado: SUCESSO\n" \
                   f"Total de teclas: {self.result['details'].get('total_keys', 0)}\n" \
                   f"Teclas pressionadas: {self.result['details'].get('pressed_keys', 0)}"
        else:
            if self.result['error']:
                return f"Teste de Teclado: FALHA\n" \
                       f"Erro: {self.result['error']}"
            else:
                return f"Teste de Teclado: FALHA\n" \
                       f"Motivo: {self.result['message']}"
    
    def cleanup(self):
        """Limpa os recursos utilizados pelo teste."""
        if self.listener:
            self.listener.stop()
        
        self.is_running = False
