"""
Módulo para coleta de informações de hardware.
Implementa a coleta de informações de diversos componentes do sistema.
"""

import os
import sys
import platform
import subprocess
import json
import re
import time

# Tenta importar bibliotecas específicas do Windows
try:
    import wmi
    import winreg
except ImportError:
    wmi = None
    winreg = None

# Tenta importar bibliotecas de terceiros
try:
    import psutil
    import cpuinfo
    import screeninfo
except ImportError:
    psutil = None
    cpuinfo = None
    screeninfo = None

class HardwareInfo:
    """Classe para coletar informações de hardware."""
    
    def __init__(self):
        """Inicializa a classe e tenta conectar ao WMI."""
        self.is_windows = platform.system() == "Windows"
        self.wmi_client = None
        
        if self.is_windows and wmi:
            try:
                self.wmi_client = wmi.WMI()
            except Exception as e:
                print(f"Erro ao conectar ao WMI: {e}")
    
    def get_all_info(self):
        """Obtém todas as informações de hardware disponíveis."""
        info = {}
        
        # Coleta informações de cada componente
        info["motherboard"] = self.get_motherboard_info()
        info["cpu"] = self.get_cpu_info()
        info["ram"] = self.get_ram_info()
        info["disks"] = self.get_disk_info()
        info["gpu"] = self.get_gpu_info()
        info["display"] = self.get_display_info()
        info["tpm"] = self.get_tpm_info()
        info["bluetooth"] = self.get_bluetooth_info()
        info["wifi"] = self.get_wifi_info()
        
        return info
    
    def _run_command(self, command, use_shell=False, timeout=10):
        """Executa comando em background sem janela visível."""
        try:
            creationflags = subprocess.CREATE_NO_WINDOW if self.is_windows else 0
            result = subprocess.run(
                command,
                shell=use_shell,
                capture_output=True,
                text=True,
                timeout=timeout,
                creationflags=creationflags
            )
            return result.stdout.strip() if result.returncode == 0 else ""
        except (subprocess.TimeoutExpired, subprocess.SubprocessError, Exception):
            return ""
    
    def _run_powershell(self, command, timeout=10):
        """Executa comando PowerShell em background."""
        return self._run_command(["powershell", "-Command", command], timeout=timeout)
    
    def get_motherboard_info(self):
        """Obtém informações da placa-mãe."""
        result = {
            "manufacturer": "Não disponível",
            "model": "Não disponível",
            "serial_number": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        # PowerShell é mais eficiente que múltiplos métodos
        try:
            # Obtém fabricante e modelo
            ps_command = "Get-WmiObject Win32_BaseBoard | Select-Object Manufacturer, Product | ConvertTo-Json"
            output = self._run_powershell(ps_command)
            
            if output:
                data = json.loads(output)
                result["manufacturer"] = data.get("Manufacturer", "Não disponível")
                result["model"] = data.get("Product", "Não disponível")
            
            # Obtém número de série da BIOS
            serial_output = self._run_powershell("(Get-CimInstance Win32_BIOS).SerialNumber")
            if serial_output:
                result["serial_number"] = serial_output
            
        except Exception as e:
            print(f"Erro ao obter informações da placa-mãe: {e}")
        
        return result
    
    def get_cpu_info(self):
        """Obtém informações do processador."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        # Prioriza cpuinfo se disponível
        if cpuinfo:
            try:
                info = cpuinfo.get_cpu_info()
                full_name = info.get("brand_raw", "")
                
                if "Intel" in full_name:
                    result["brand"] = "Intel"
                    result["model"] = full_name.replace("Intel", "").strip()
                elif "AMD" in full_name:
                    result["brand"] = "AMD"
                    result["model"] = full_name.replace("AMD", "").strip()
                else:
                    result["model"] = full_name
                return result
            except Exception:
                pass
        
        # Fallback para WMI/Windows
        if self.is_windows:
            try:
                output = self._run_command(["wmic", "cpu", "get", "name"])
                lines = output.split('\n')
                if len(lines) > 1:
                    full_name = lines[1].strip()
                    if "Intel" in full_name:
                        result["brand"] = "Intel"
                        result["model"] = full_name.replace("Intel", "").strip()
                    elif "AMD" in full_name:
                        result["brand"] = "AMD"
                        result["model"] = full_name.replace("AMD", "").strip()
                    else:
                        result["model"] = full_name
            except Exception as e:
                print(f"Erro ao obter informações do CPU: {e}")
        
        return result
    
    def get_ram_info(self):
        """Obtém informações da memória RAM."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        if not self.is_windows:
            return result
        
        try:
            ps_command = "Get-CimInstance Win32_PhysicalMemory | Select-Object BankLabel, Capacity, Speed | ConvertTo-Json"
            output = self._run_powershell(ps_command, timeout=15)
            
            if output:
                data = json.loads(output)
                modules = []
                total_memory_gb = 0
                
                if isinstance(data, list):
                    for module in data:
                        capacity = module.get("Capacity", 0)
                        size_gb = round(int(capacity) / (1024**3), 2) if capacity else 0
                        total_memory_gb += size_gb
                        
                        modules.append({
                            "banklabel": module.get("BankLabel", "BANK N/A"),
                            "size": f"{size_gb:.2f} GB",
                            "speed": f"{module.get('Speed', 0)} MHz"
                        })
                
                if total_memory_gb > 0:
                    result["total"] = f"{total_memory_gb:.2f} GB"
                    result["slots_used"] = str(len(modules))
                    result["modules"] = modules
                    
        except Exception as e:
            print(f"Erro ao obter informações da RAM: {e}")
        
        return result
    
    def get_disk_info(self):
        """Obtém informações dos discos."""
        result = []
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações dos discos físicos
            output = self._run_command(["wmic", "diskdrive", "get", "model,size,mediatype"])
            
            # Processa a saída
            lines = output.strip().split("\n")
            if len(lines) < 2:
                return result
                
            header = lines[0]
            
            # Determina os índices das colunas
            model_index = header.lower().find("model")
            size_index = header.lower().find("size")
            mediatype_index = header.lower().find("mediatype")
            
            # Ordena os índices para extrair corretamente
            indices = sorted([ 
                (model_index, "model"), 
                (size_index, "size"),
                (mediatype_index, "mediatype")
            ])
            
            for line in lines[1:]:
                if not line.strip():
                    continue
                
                disk_data = {}
                for i, (index, key) in enumerate(indices):
                    if index < 0:
                        continue
                    
                    # Calcula o fim da coluna atual
                    if i + 1 < len(indices):
                        next_index = indices[i+1][0]
                        end_index = next_index if next_index > index else len(line)
                    else:
                        end_index = len(line)
                    
                    # Extrai o valor
                    value = line[index:end_index].strip()
                    disk_data[key] = value
                
                try:
                    # Converte o tamanho para GB
                    size_gb = round(int(disk_data.get("size", 0)) / (1024**3), 2)
                    # Adiciona o disco à lista
                    result.append({
                        "model": disk_data.get("model", "Não disponível"),
                        "size": f"{size_gb:.2f} GB",
                        "type": disk_data.get("mediatype", "Não disponível")
                    })
                except (ValueError, IndexError):
                    continue
        except Exception as e:
            print(f"Erro ao obter informações dos discos: {e}")
        
        return result
    
    def get_gpu_info(self):
        """Obtém informações da placa de vídeo."""
        result = []
        
        if not self.is_windows:
            return result
        
        try:
            output = self._run_command(["wmic", "path", "Win32_VideoController", "get", "name"])
            lines = output.strip().split('\n')
            
            for line in lines[1:]:
                if line.strip():
                    result.append({"model": line.strip()})
        except Exception as e:
            print(f"Erro ao obter informações da GPU: {e}")
        
        return result
    
    def get_display_info(self):
        """Obtém informações do display."""
        result = {"resolution": "Não disponível"}
        
        if screeninfo:
            try:
                monitors = screeninfo.get_monitors()
                if monitors:
                    primary = next((m for m in monitors if m.is_primary), monitors[0])
                    result["resolution"] = f"{primary.width}x{primary.height}"
            except Exception as e:
                print(f"Erro ao obter informações do display: {e}")
        
        return result
    
    def get_tpm_info(self):
        """Obtém informações do TPM."""
        result = {
            "version": "Não disponível",
            "status": "Não disponível",
            "manufacturer": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            ps_command = "Get-Tpm | Select-Object TpmPresent,ManufacturerVersion,ManufacturerIdTxt | ConvertTo-Json"
            output = self._run_powershell(ps_command)
            
            if output:
                data = json.loads(output)
                if data.get("TpmPresent"):
                    result["status"] = "Presente"
                    result["version"] = data.get("ManufacturerVersion", "Não disponível")
                    result["manufacturer"] = data.get("ManufacturerIdTxt", "Não disponível")
                else:
                    result["status"] = "Não presente"
        except Exception as e:
            print(f"Erro ao obter informações do TPM: {e}")
        
        return result
    
    def get_bluetooth_info(self):
        """Obtém informações do Bluetooth."""
        result = {
            "device_name": "Não disponível",
            "device_status": "Não disponível"
        }

        if not self.is_windows:
            return result

        try:
            # Status do serviço Bluetooth
            status_output = self._run_powershell('Get-Service bthserv | Select-Object -ExpandProperty Status')
            if status_output.lower() == "running":
                result["device_status"] = "Ativo"
            else:
                result["device_status"] = status_output.capitalize()

            # Fabricante do dispositivo Bluetooth
            device_output = self._run_powershell('Get-PnpDevice | Where-Object {$_.Class -like "Bluetooth"} | Select-Object -First 1 -ExpandProperty Manufacturer')
            if device_output and device_output.lower() != "microsoft":
                result["device_name"] = device_output
        except Exception as e:
            print(f"Erro ao obter informações do Bluetooth: {e}")

        return result
    
    def get_wifi_info(self):
        """Obtém informações do Wi-Fi."""
        result = {
            "adapter_name": "Não disponível",
            "adapter_status": "Não disponível",
            "connected_ssid": "Não disponível"
        }

        if not self.is_windows:
            return result

        try:
            # Nome do adaptador Wi-Fi
            adapter_output = self._run_powershell('Get-NetAdapter | Where-Object {$_.InterfaceDescription -match "Wi|Wireless|802"} | Select-Object -First 1 -ExpandProperty InterfaceDescription')
            if adapter_output:
                result["adapter_name"] = adapter_output

            # Status do adaptador
            status_output = self._run_powershell('Get-NetAdapter | Where-Object {$_.InterfaceDescription -match "Wi|Wireless|802"} | Select-Object -First 1 -ExpandProperty Status')
            if status_output:
                result["adapter_status"] = status_output

            # SSID conectado
            ssid_output = self._run_powershell('(netsh wlan show interfaces) -match "^ *SSID" | Select-Object -First 1 | ForEach-Object { ($_ -split ":")[1].Trim() }')
            if ssid_output:
                result["connected_ssid"] = ssid_output
        except Exception as e:
            print(f"Erro ao obter informações do Wi-Fi: {e}")

        return result