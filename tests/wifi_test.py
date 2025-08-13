"""
Módulo para teste de Wi-Fi.
Implementação real que interage diretamente com o hardware.
"""

import os
import time
import subprocess
import re
import platform
import ctypes
from datetime import datetime

class WiFiTest:
    """Classe para teste de Wi-Fi."""
    
    def __init__(self):
        """Inicializa o teste de Wi-Fi."""
        self.result = {
            'success': False,
            'execution_time': 0,
            'adapter_present': False,
            'adapter_name': 'Não disponível',
            'adapter_status': 'Não disponível',
            'connected_ssid': 'Não disponível',
            'signal_strength': 'Não disponível',
            'ip_address': 'Não disponível',
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
        """Executa o teste de Wi-Fi."""
        start_time = time.time()
        
        try:
            # Método 1: Usando netsh
            if self._check_wifi_netsh():
                self.result['success'] = True
            # Método 2: Usando WMI
            elif self._check_wifi_wmi():
                self.result['success'] = True
            # Método 3: Usando PowerShell
            elif self._check_wifi_powershell():
                self.result['success'] = True
            else:
                self.result['error'] = "Não foi possível detectar adaptador Wi-Fi."
                self.result['success'] = False
        except Exception as e:
            self.result['error'] = f"Erro ao executar teste de Wi-Fi: {e}"
            self.result['success'] = False
        
        self.result['execution_time'] = time.time() - start_time
        return self.result['success']
    
    def _check_wifi_netsh(self):
        """Verifica o Wi-Fi usando netsh."""
        try:
            # Executa o comando para listar interfaces de rede sem fio
            result = subprocess.check_output(
                ["netsh", "wlan", "show", "interfaces"],
                universal_newlines=True
            )
            
            # Verifica se há adaptadores Wi-Fi
            if "Não há nenhuma interface sem fio no sistema" in result or "There is no wireless interface on the system" in result:
                return False
            
            # Extrai informações do adaptador
            name_match = re.search(r"Nome\s*:\s*(.*?)[\r\n]", result) or re.search(r"Name\s*:\s*(.*?)[\r\n]", result)
            state_match = re.search(r"Estado\s*:\s*(.*?)[\r\n]", result) or re.search(r"State\s*:\s*(.*?)[\r\n]", result)
            ssid_match = re.search(r"SSID\s*:\s*(.*?)[\r\n]", result)
            signal_match = re.search(r"Sinal\s*:\s*(.*?)[\r\n]", result) or re.search(r"Signal\s*:\s*(.*?)[\r\n]", result)
            
            if name_match:
                self.result['adapter_present'] = True
                self.result['adapter_name'] = name_match.group(1).strip()
                
                if state_match:
                    self.result['adapter_status'] = state_match.group(1).strip()
                
                if ssid_match:
                    self.result['connected_ssid'] = ssid_match.group(1).strip()
                
                if signal_match:
                    self.result['signal_strength'] = signal_match.group(1).strip()
                
                # Obtém o endereço IP
                self._get_ip_address()
                
                return True
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Wi-Fi via netsh: {e}")
            return False
    
    def _check_wifi_wmi(self):
        """Verifica o Wi-Fi usando WMI."""
        try:
            import wmi
            wmi_client = wmi.WMI()
            
            # Verifica adaptadores de rede sem fio
            network_adapters = wmi_client.Win32_NetworkAdapter(
                AdapterTypeId=9  # Wireless
            )
            
            if not network_adapters:
                # Tenta outro método para identificar adaptadores Wi-Fi
                network_adapters = wmi_client.Win32_NetworkAdapter(
                    ConfigManagerErrorCode=0
                )
                
                network_adapters = [
                    adapter for adapter in network_adapters
                    if "wireless" in adapter.Name.lower() or "wi-fi" in adapter.Name.lower() or "wifi" in adapter.Name.lower()
                ]
            
            if network_adapters:
                adapter = network_adapters[0]
                self.result['adapter_present'] = True
                self.result['adapter_name'] = adapter.Name
                self.result['adapter_status'] = 'Conectado' if adapter.NetConnectionStatus == 2 else 'Desconectado'
                
                # Tenta obter informações de conexão
                self._get_wifi_connection_info()
                
                # Obtém o endereço IP
                self._get_ip_address()
                
                return True
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Wi-Fi via WMI: {e}")
            return False
    
    def _check_wifi_powershell(self):
        """Verifica o Wi-Fi usando PowerShell."""
        try:
            # Comando PowerShell para obter informações do adaptador Wi-Fi
            ps_command = """
            $wifiAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like '*Wireless*' -or $_.InterfaceDescription -like '*Wi-Fi*' -or $_.InterfaceDescription -like '*WiFi*' }
            if ($wifiAdapters) {
                $adapter = $wifiAdapters[0]
                $wifiInfo = @{
                    AdapterPresent = $true
                    AdapterName = $adapter.InterfaceDescription
                    AdapterStatus = $adapter.Status
                    MacAddress = $adapter.MacAddress
                }
                
                # Tenta obter informações da rede conectada
                try {
                    $wifiNetworks = (netsh wlan show interfaces) -join "`n"
                    $ssidMatch = [regex]::Match($wifiNetworks, "SSID\s*:\s*(.*?)[\r\n]")
                    $signalMatch = [regex]::Match($wifiNetworks, "Signal\s*:\s*(.*?)[\r\n]")
                    
                    if ($ssidMatch.Success) {
                        $wifiInfo.ConnectedSSID = $ssidMatch.Groups[1].Value.Trim()
                    }
                    
                    if ($signalMatch.Success) {
                        $wifiInfo.SignalStrength = $signalMatch.Groups[1].Value.Trim()
                    }
                } catch {}
                
                # Tenta obter o endereço IP
                try {
                    $ipConfig = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4
                    if ($ipConfig) {
                        $wifiInfo.IPAddress = $ipConfig.IPAddress
                    }
                } catch {}
                
                $wifiInfo | ConvertTo-Json
            } else {
                @{ AdapterPresent = $false } | ConvertTo-Json
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
                    wifi_info = json.loads(stdout)
                    
                    if wifi_info.get('AdapterPresent', False):
                        self.result['adapter_present'] = True
                        self.result['adapter_name'] = wifi_info.get('AdapterName', 'Não disponível')
                        self.result['adapter_status'] = wifi_info.get('AdapterStatus', 'Não disponível')
                        self.result['connected_ssid'] = wifi_info.get('ConnectedSSID', 'Não disponível')
                        self.result['signal_strength'] = wifi_info.get('SignalStrength', 'Não disponível')
                        self.result['ip_address'] = wifi_info.get('IPAddress', 'Não disponível')
                        
                        return True
                except json.JSONDecodeError:
                    pass
            
            return False
        except Exception as e:
            print(f"Erro ao verificar Wi-Fi via PowerShell: {e}")
            return False
    
    def _get_wifi_connection_info(self):
        """Obtém informações da conexão Wi-Fi atual."""
        try:
            # Executa o comando para obter informações da rede sem fio
            result = subprocess.check_output(
                ["netsh", "wlan", "show", "interfaces"],
                universal_newlines=True
            )
            
            # Extrai informações da conexão
            ssid_match = re.search(r"SSID\s*:\s*(.*?)[\r\n]", result)
            signal_match = re.search(r"Sinal\s*:\s*(.*?)[\r\n]", result) or re.search(r"Signal\s*:\s*(.*?)[\r\n]", result)
            
            if ssid_match:
                self.result['connected_ssid'] = ssid_match.group(1).strip()
            
            if signal_match:
                self.result['signal_strength'] = signal_match.group(1).strip()
        except Exception as e:
            print(f"Erro ao obter informações da conexão Wi-Fi: {e}")
    
    def _get_ip_address(self):
        """Obtém o endereço IP do adaptador Wi-Fi."""
        try:
            # Executa o comando para obter informações de IP
            result = subprocess.check_output(
                ["ipconfig"],
                universal_newlines=True
            )
            
            # Procura por seções relacionadas a Wi-Fi ou Wireless
            wifi_sections = re.findall(r"(Wi-Fi|Wireless|Sem Fio).*?IPv4.*?:(.*?)[\r\n]", result, re.DOTALL | re.IGNORECASE)
            
            if wifi_sections:
                for section in wifi_sections:
                    ip_match = re.search(r"IPv4.*?:\s*([\d\.]+)", section[1])
                    if ip_match:
                        self.result['ip_address'] = ip_match.group(1).strip()
                        break
        except Exception as e:
            print(f"Erro ao obter endereço IP do Wi-Fi: {e}")
    
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
            f"Teste de Wi-Fi: {status}\n"
            f"Tempo de execução: {self.result['execution_time']:.2f} segundos\n"
            f"Adaptador presente: {'Sim' if self.result['adapter_present'] else 'Não'}\n"
            f"Nome/Modelo: {self.result['adapter_name']}\n"
            f"Status: {self.result['adapter_status']}\n"
            f"SSID Conectado: {self.result['connected_ssid']}\n"
            f"Intensidade do Sinal: {self.result['signal_strength']}\n"
            f"Endereço IP: {self.result['ip_address']}"
        )


# Função para teste rápido
if __name__ == "__main__":
    wifi_test = WiFiTest()
    
    if wifi_test.initialize():
        print("Executando teste de Wi-Fi...")
        wifi_test.execute()
        print(wifi_test.get_formatted_result())
    else:
        print(f"Não foi possível inicializar o teste: {wifi_test.result['error']}")
