"""
Módulo para teste de Bluetooth.
Implementação real que interage diretamente com o hardware.
"""

import os
import time
import subprocess
import re
import platform
import ctypes
from datetime import datetime

class BluetoothTest:
    """Classe para teste de Bluetooth."""
    
    def __init__(self):
        """Inicializa o teste de Bluetooth."""
        self.result = {
            'success': False,
            'execution_time': 0,
            'device_present': False,
            'device_name': 'Não disponível',
            'device_status': 'Não disponível',
            'device_address': 'Não disponível',
            'error': None
        }
        self.is_windows = platform.system() == 'Windows'
    
    def initialize(self):
        """Inicializa o teste, verificando se o sistema é compatível."""
        if not self.is_windows:
            self.result['error'] = "Este teste só é compatível com Windows."
            return False
        
        # Verifica se está sendo executado como administrador
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                self.result['error'] = "Este teste requer privilégios de administrador."
                return False
        except Exception as e:
            self.result['error'] = f"Erro ao verificar privilégios: {e}"
            return False
        
        return True
    
    def execute(self):
        """Executa o teste de Bluetooth."""
        start_time = time.time()
        
        try:
            # Método 1: Usando WMI
            if self._check_bluetooth_wmi():
                self.result['success'] = True
            # Método 2: Usando PowerShell
            elif self._check_bluetooth_powershell():
                self.result['success'] = True
            # Método 3: Usando comandos do sistema
            elif self._check_bluetooth_system():
                self.result['success'] = True
            else:
                self.result['error'] = "Não foi possível detectar dispositivo Bluetooth."
                self.result['success'] = False
        except Exception as e:
            self.result['error'] = f"Erro ao executar teste de Bluetooth: {e}"
            self.result['success'] = False
        
        self.result['execution_time'] = time.time() - start_time
        return self.result['success']
    
    def _check_bluetooth_wmi(self):
        """Verifica o Bluetooth usando WMI."""
        try:
            import wmi
            wmi_client = wmi.WMI()
            
            # Verifica adaptadores Bluetooth
            bluetooth_devices = wmi_client.Win32_PnPEntity(
                ConfigManagerErrorCode=0,
                Service="BTHUSB"
            )
            
            if bluetooth_devices:
                self.result['device_present'] = True
                self.result['device_name'] = bluetooth_devices[0].Name
                self.result['device_status'] = 'Ativado' if bluetooth_devices[0].Status == 'OK' else bluetooth_devices[0].Status
                
                # Tenta obter o endereço MAC
                self._get_bluetooth_address()
                
                return True
            
            # Método alternativo para detectar dispositivos Bluetooth
            bluetooth_devices = wmi_client.Win32_PnPEntity(
                ConfigManagerErrorCode=0,
                Description="Bluetooth"
            )
            
            if bluetooth_devices:
                self.result['device_present'] = True
                self.result['device_name'] = bluetooth_devices[0].Name
                self.result['device_status'] = 'Ativado' if bluetooth_devices[0].Status == 'OK' else bluetooth_devices[0].Status
                
                # Tenta obter o endereço MAC
                self._get_bluetooth_address()
                
                return True
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Bluetooth via WMI: {e}")
            return False
    
    def _check_bluetooth_powershell(self):
        """Verifica o Bluetooth usando PowerShell."""
        try:
            # Comando PowerShell para obter informações do Bluetooth
            ps_command = """
            $bluetoothRadios = Get-PnpDevice | Where-Object { $_.Class -eq 'Bluetooth' -or $_.FriendlyName -like '*Bluetooth*' }
            if ($bluetoothRadios) {
                $bluetoothInfo = @{
                    DevicePresent = $true
                    DeviceName = $bluetoothRadios[0].FriendlyName
                    DeviceStatus = $bluetoothRadios[0].Status
                    DeviceID = $bluetoothRadios[0].InstanceId
                }
                $bluetoothInfo | ConvertTo-Json
            } else {
                @{ DevicePresent = $false } | ConvertTo-Json
            }
            """
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            stdout, stderr = process.communicate()
            
            if process.returncode == 0 and stdout.strip():
                import json
                try:
                    bluetooth_info = json.loads(stdout)
                    
                    if bluetooth_info.get('DevicePresent', False):
                        self.result['device_present'] = True
                        self.result['device_name'] = bluetooth_info.get('DeviceName', 'Não disponível')
                        self.result['device_status'] = bluetooth_info.get('DeviceStatus', 'Não disponível')
                        
                        # Tenta obter o endereço MAC
                        self._get_bluetooth_address()
                        
                        return True
                except json.JSONDecodeError:
                    pass
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Bluetooth via PowerShell: {e}")
            return False
    
    def _check_bluetooth_system(self):
        """Verifica o Bluetooth usando comandos do sistema."""
        try:
            # Executa o comando para listar dispositivos Bluetooth
            result = subprocess.check_output(
                ["wmic", "path", "Win32_PnPEntity", "where", "Service='BTHUSB'", "get", "Name,Status", "/format:csv"],
                universal_newlines=True
            )
            
            lines = result.strip().split('\n')
            if len(lines) > 1:  # Primeira linha é o cabeçalho
                parts = lines[1].split(',')
                if len(parts) >= 3:
                    self.result['device_present'] = True
                    self.result['device_name'] = parts[1].strip()
                    self.result['device_status'] = 'Ativado' if parts[2].strip() == 'OK' else parts[2].strip()
                    
                    # Tenta obter o endereço MAC
                    self._get_bluetooth_address()
                    
                    return True
            
            # Método alternativo
            result = subprocess.check_output(
                ["wmic", "path", "Win32_PnPEntity", "where", "Description like '%Bluetooth%'", "get", "Name,Status", "/format:csv"],
                universal_newlines=True
            )
            
            lines = result.strip().split('\n')
            if len(lines) > 1:  # Primeira linha é o cabeçalho
                parts = lines[1].split(',')
                if len(parts) >= 3:
                    self.result['device_present'] = True
                    self.result['device_name'] = parts[1].strip()
                    self.result['device_status'] = 'Ativado' if parts[2].strip() == 'OK' else parts[2].strip()
                    
                    # Tenta obter o endereço MAC
                    self._get_bluetooth_address()
                    
                    return True
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Bluetooth via comandos do sistema: {e}")
            return False
    
    def _get_bluetooth_address(self):
        """Obtém o endereço MAC do adaptador Bluetooth."""
        try:
            # Comando PowerShell para obter o endereço MAC do adaptador Bluetooth
            ps_command = """
            $bluetoothRadios = Get-PnpDevice | Where-Object { $_.Class -eq 'Bluetooth' -or $_.FriendlyName -like '*Bluetooth*' }
            if ($bluetoothRadios) {
                $deviceId = $bluetoothRadios[0].InstanceId
                $address = (Get-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\$deviceId\\Device Parameters" -Name "Address" -ErrorAction SilentlyContinue).Address
                if ($address) {
                    $address = [BitConverter]::ToString([byte[]]$address) -replace '-', ':'
                    $address
                }
            }
            """
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            stdout, stderr = process.communicate()
            
            if process.returncode == 0 and stdout.strip():
                self.result['device_address'] = stdout.strip()
        except Exception as e:
            print(f"Erro ao obter endereço MAC do Bluetooth: {e}")
    
    def get_result(self):
        """Retorna o resultado do teste."""
        return self.result
    
    def cleanup(self):
        """Limpa recursos utilizados pelo teste."""
        # Não há recursos a serem limpos neste teste
        pass
    
    def get_formatted_result(self):
        """Retorna o resultado formatado para exibição."""
        if self.result['error']:
            return f"Erro: {self.result['error']}"
        
        status = "SUCESSO" if self.result['success'] else "FALHA"
        
        return (
            f"Teste de Bluetooth: {status}\n"
            f"Tempo de execução: {self.result['execution_time']:.2f} segundos\n"
            f"Dispositivo presente: {'Sim' if self.result['device_present'] else 'Não'}\n"
            f"Nome/Modelo: {self.result['device_name']}\n"
            f"Status: {self.result['device_status']}\n"
            f"Endereço MAC: {self.result['device_address']}"
        )


# Função para teste rápido
if __name__ == "__main__":
    bluetooth_test = BluetoothTest()
    
    if bluetooth_test.initialize():
        print("Executando teste de Bluetooth...")
        bluetooth_test.execute()
        print(bluetooth_test.get_formatted_result())
    else:
        print(f"Não foi possível inicializar o teste: {bluetooth_test.result['error']}")
