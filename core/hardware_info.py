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
        
        # Verifica se as bibliotecas necessárias estão disponíveis
        if not psutil:
            print("Aviso: Biblioteca psutil não encontrada. Algumas informações podem não estar disponíveis.")
        if not cpuinfo:
            print("Aviso: Biblioteca py-cpuinfo não encontrada. Algumas informações podem não estar disponíveis.")
        if not screeninfo:
            print("Aviso: Biblioteca screeninfo não encontrada. Algumas informações podem não estar disponíveis.")
        if self.is_windows and not wmi:
            print("Aviso: Biblioteca WMI não encontrada. Algumas informações podem não estar disponíveis no Windows.")
        if self.is_windows and not winreg:
            print("Aviso: Módulo winreg não encontrado. Algumas informações podem não estar disponíveis no Windows.")
    
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
    
    def get_motherboard_info(self):
        """Obtém informações da placa-mãe."""
        result = {
            "manufacturer": "Não disponível",
            "model": "Não disponível",
            "serial_number": "Não disponível"
        }
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_motherboard_info_wmi,
            self._get_motherboard_info_wmic,
            self._get_motherboard_info_registry
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores foram preenchidos, interrompe a busca
                    if all(value != "Não disponível" for value in result.values()):
                        break
            except Exception as e:
                print(f"Erro ao obter informações da placa-mãe usando {method.__name__}: {e}")
        
        return result
    
    def _get_motherboard_info_wmi(self):
        """Obtém informações da placa-mãe usando WMI."""
        result = {
            "manufacturer": "Não disponível",
            "model": "Não disponível",
            "serial_number": "Não disponível"
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            for board in self.wmi_client.Win32_BaseBoard():
                if hasattr(board, "Manufacturer") and board.Manufacturer:
                    result["manufacturer"] = board.Manufacturer.strip()
                if hasattr(board, "Product") and board.Product:
                    result["model"] = board.Product.strip()
                if hasattr(board, "SerialNumber") and board.SerialNumber:
                    result["serial_number"] = board.SerialNumber.strip()
                
                # Se encontrou informações, interrompe o loop
                if result["model"] != "Não disponível":
                    break
        except Exception as e:
            print(f"Erro ao obter informações da placa-mãe via WMI: {e}")
        
        return result
    
    def _get_motherboard_info_wmic(self):
        """Obtém informações da placa-mãe usando WMIC."""
        result = {
            "manufacturer": "Não disponível",
            "model": "Não disponível",
            "serial_number": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações da placa-mãe
            output = subprocess.check_output(
                ["wmic", "baseboard", "get", "manufacturer,product,serialnumber"],
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            lines = output.strip().split("\n")
            if len(lines) > 1:  # Pelo menos uma linha além do cabeçalho
                # Determina os índices das colunas
                header = lines[0]
                manufacturer_index = header.lower().find("manufacturer")
                product_index = header.lower().find("product")
                serial_index = header.lower().find("serialnumber")
                
                # Ordena os índices para extrair corretamente
                indices = sorted([(manufacturer_index, "manufacturer"), (product_index, "model"), (serial_index, "serial_number")])
                
                line = lines[1].strip()
                last_index = 0
                for i, (index, key) in enumerate(indices):
                    if index < 0: continue
                    
                    # Calcula o fim da coluna atual
                    end_index = indices[i+1][0] if i + 1 < len(indices) and indices[i+1][0] >= 0 else len(line)
                    
                    # Extrai o valor
                    value = line[index:end_index].strip()
                    if value:
                        result[key] = value
        except Exception as e:
            print(f"Erro ao obter informações da placa-mãe via WMIC: {e}")
        
        return result
    
    def _get_motherboard_info_registry(self):
        """Obtém informações da placa-mãe usando o registro do Windows."""
        result = {
            "manufacturer": "Não disponível",
            "model": "Não disponível",
            "serial_number": "Não disponível" # Número de série geralmente não está no registro
        }
        
        if not self.is_windows or not winreg:
            return result
        
        try:
            # Comando para obter informações do registro
            reg_command = (
                "reg query \"HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\BIOS\" "
                "/v BaseBoardManufacturer /v BaseBoardProduct"
            )
            
            output = subprocess.check_output(
                reg_command,
                shell=True,
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            for line in output.strip().split("\n"):
                if "BaseBoardManufacturer" in line and "REG_SZ" in line:
                    parts = line.split("REG_SZ")
                    if len(parts) > 1:
                        result["manufacturer"] = parts[1].strip()
                elif "BaseBoardProduct" in line and "REG_SZ" in line:
                    parts = line.split("REG_SZ")
                    if len(parts) > 1:
                        result["model"] = parts[1].strip()
        except Exception as e:
            print(f"Erro ao obter informações da placa-mãe via Registro: {e}")
        
        return result
    
    def get_cpu_info(self):
        """Obtém informações do processador."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_cpu_info_cpuinfo,
            self._get_cpu_info_wmi,
            self._get_cpu_info_wmic,
            self._get_cpu_info_registry
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores foram preenchidos, interrompe a busca
                    if all(value != "Não disponível" for value in result.values()):
                        break
            except Exception as e:
                print(f"Erro ao obter informações do processador usando {method.__name__}: {e}")
        
        return result
    
    def _get_cpu_info_cpuinfo(self):
        """Obtém informações do processador usando a biblioteca cpuinfo."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        if not cpuinfo:
            return result
            
        try:
            info = cpuinfo.get_cpu_info()
            
            if "brand_raw" in info and info["brand_raw"]:
                full_name = info["brand_raw"]
                
                # Extrai marca e modelo do nome completo
                if "Intel" in full_name:
                    result["brand"] = "Intel"
                    result["model"] = full_name.replace("Intel", "").strip()
                elif "AMD" in full_name:
                    result["brand"] = "AMD"
                    result["model"] = full_name.replace("AMD", "").strip()
                else:
                    # Se não conseguir identificar a marca, usa o nome completo como modelo
                    result["model"] = full_name
        except Exception as e:
            print(f"Erro ao obter informações do processador via cpuinfo: {e}")
        
        return result
    
    def _get_cpu_info_wmi(self):
        """Obtém informações do processador usando WMI."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            for processor in self.wmi_client.Win32_Processor():
                full_name = processor.Name if hasattr(processor, "Name") and processor.Name else ""
                
                # Extrai marca e modelo do nome completo
                if "Intel" in full_name:
                    result["brand"] = "Intel"
                    result["model"] = full_name.replace("Intel", "").strip()
                elif "AMD" in full_name:
                    result["brand"] = "AMD"
                    result["model"] = full_name.replace("AMD", "").strip()
                else:
                    # Se não conseguir identificar a marca, usa o nome completo como modelo
                    result["model"] = full_name
                
                # Se encontrou informações, interrompe o loop
                if result["model"] != "Não disponível":
                    break
        except Exception as e:
            print(f"Erro ao obter informações do processador via WMI: {e}")
        
        return result
    
    def _get_cpu_info_wmic(self):
        """Obtém informações do processador usando WMIC."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter o nome do processador
            output = subprocess.check_output(
                ["wmic", "cpu", "get", "name"],
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            lines = output.strip().split("\n")
            if len(lines) > 1:
                full_name = lines[1].strip()
                
                # Extrai marca e modelo do nome completo
                if "Intel" in full_name:
                    result["brand"] = "Intel"
                    result["model"] = full_name.replace("Intel", "").strip()
                elif "AMD" in full_name:
                    result["brand"] = "AMD"
                    result["model"] = full_name.replace("AMD", "").strip()
                else:
                    # Se não conseguir identificar a marca, usa o nome completo como modelo
                    result["model"] = full_name
        except Exception as e:
            print(f"Erro ao obter informações do processador via WMIC: {e}")
        
        return result
    
    def _get_cpu_info_registry(self):
        """Obtém informações do processador usando o registro do Windows."""
        result = {
            "brand": "Não disponível",
            "model": "Não disponível"
        }
        
        if not self.is_windows or not winreg:
            return result
        
        try:
            # Comando para obter informações do registro
            reg_command = (
                "reg query \"HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0\" "
                "/v ProcessorNameString"
            )
            
            output = subprocess.check_output(
                reg_command,
                shell=True,
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            for line in output.strip().split("\n"):
                if "ProcessorNameString" in line and "REG_SZ" in line:
                    parts = line.split("REG_SZ")
                    if len(parts) > 1:
                        full_name = parts[1].strip()
                        
                        # Extrai marca e modelo do nome completo
                        if "Intel" in full_name:
                            result["brand"] = "Intel"
                            result["model"] = full_name.replace("Intel", "").strip()
                        elif "AMD" in full_name:
                            result["brand"] = "AMD"
                            result["model"] = full_name.replace("AMD", "").strip()
                        else:
                            # Se não conseguir identificar a marca, usa o nome completo como modelo
                            result["model"] = full_name
        except Exception as e:
            print(f"Erro ao obter informações do processador via Registro: {e}")
        
        return result
    
    def get_ram_info(self):
        """Obtém informações da memória RAM."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        # Lista de métodos para tentar obter as informações, priorizando PowerShell CIM
        methods = [
            self._get_ram_info_powershell_cim, # Novo método prioritário
            self._get_ram_info_wmi,
            self._get_ram_info_wmic,
            self._get_ram_info_psutil
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if key == "modules":
                            if value:  # Se há módulos, atualiza a lista
                                result["modules"] = value
                        elif value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores principais foram preenchidos, interrompe a busca
                    if result["total"] != "Não disponível" and result["slots_used"] != "Não disponível" and result["modules"]:
                        break
            except Exception as e:
                print(f"Erro ao obter informações da memória RAM usando {method.__name__}: {e}")
        
        return result
    
    def _get_ram_info_powershell_cim(self):
        """Obtém informações da memória RAM usando PowerShell com CIM."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Comando PowerShell para obter informações detalhadas da memória
            ps_command = "Get-CimInstance -ClassName Win32_PhysicalMemory | " \
                         "Select-Object BankLabel, DeviceLocator, Capacity, Manufacturer, " \
                         "PartNumber, SerialNumber, Speed | ConvertTo-Json"
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW # Evita abrir janela do PowerShell
            )
            stdout, stderr = process.communicate(timeout=15) # Aumenta timeout
            
            if process.returncode == 0 and stdout.strip():
                try:
                    # Tenta processar como JSON
                    data = json.loads(stdout)
                    
                    # Converte para lista se for um único objeto
                    if not isinstance(data, list):
                        data = [data]
                    
                    total_memory_gb = 0
                    modules = []
                    
                    for module in data:
                        # Calcula o tamanho em GB
                        size_gb = round(int(module.get("Capacity", 0)) / (1024**3), 2)
                        total_memory_gb += size_gb
                        
                        # Obtém o fabricante
                        manufacturer = module.get("Manufacturer", "").strip()
                        
                        # Tenta obter mais informações se o fabricante estiver vazio ou for genérico
                        if not manufacturer or manufacturer.lower() in ["unknown", "not specified", "0000", "to be filled by o.e.m."]:
                            part_number = module.get("PartNumber", "").strip()
                            if part_number:
                                # Tenta identificar o fabricante pelo part number
                                if "KHX" in part_number or "HX" in part_number:
                                    manufacturer = "Kingston"
                                elif "CMK" in part_number or "CMW" in part_number or "CM" in part_number:
                                    manufacturer = "Corsair"
                                elif "F4" in part_number and "G" in part_number:
                                    manufacturer = "G.Skill"
                                elif "BLS" in part_number or "CT" in part_number:
                                    manufacturer = "Crucial"
                                elif "AX4U" in part_number:
                                    manufacturer = "ADATA XPG"
                                elif "M378" in part_number:
                                    manufacturer = "Samsung"
                                elif "HMA" in part_number or "HMP" in part_number:
                                    manufacturer = "Hynix"
                                elif "PVS" in part_number:
                                    manufacturer = "Patriot Viper"
                                elif "TED4" in part_number:
                                    manufacturer = "Teamgroup Elite"
                                else:
                                    manufacturer = "Não identificado" # Mantém como não identificado se não reconhecer
                        
                        # Se ainda estiver vazio, usa "Não identificado"
                        if not manufacturer or manufacturer.lower() in ["unknown", "not specified", "0000", "to be filled by o.e.m."]:
                            manufacturer = "Não identificado"
                        
                        # Adiciona o módulo à lista
                        bank_label = module.get("BankLabel", "").strip() or module.get("DeviceLocator", "").strip()
                        modules.append({
                            "size": f"{size_gb} GB",
                            "manufacturer": manufacturer,
                            "location": bank_label,
                            "speed": f"{module.get("Speed", 0)} MHz"
                        })
                    
                    # Atualiza o resultado
                    if total_memory_gb > 0:
                        result["total"] = f"{total_memory_gb:.2f} GB"
                    
                    if modules:
                        result["slots_used"] = str(len(modules))
                        result["modules"] = modules
                except json.JSONDecodeError:
                    print("Erro ao decodificar JSON da saída do PowerShell para RAM")
        except Exception as e:
            print(f"Erro ao obter informações da memória RAM via PowerShell CIM: {e}")
        
        return result

    def _get_ram_info_wmi(self):
        """Obtém informações da memória RAM usando WMI."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Obtém informações dos módulos de memória
            physical_memory = self.wmi_client.Win32_PhysicalMemory()
            
            total_memory_gb = 0
            modules = []
            
            for memory in physical_memory:
                # Calcula o tamanho em GB
                size_gb = round(int(memory.Capacity) / (1024**3), 2) if hasattr(memory, "Capacity") and memory.Capacity else 0
                total_memory_gb += size_gb
                
                # Obtém o fabricante
                manufacturer = memory.Manufacturer if hasattr(memory, "Manufacturer") and memory.Manufacturer else "Não disponível"
                if manufacturer.lower() in ["unknown", "not specified", "0000", "to be filled by o.e.m."]:
                     manufacturer = "Não identificado"
                
                # Adiciona o módulo à lista
                bank_label = memory.BankLabel if hasattr(memory, "BankLabel") and memory.BankLabel else "Não disponível"
                device_locator = memory.DeviceLocator if hasattr(memory, "DeviceLocator") and memory.DeviceLocator else "Não disponível"
                location = bank_label if bank_label != "Não disponível" else device_locator
                
                modules.append({
                    "size": f"{size_gb} GB",
                    "manufacturer": manufacturer,
                    "location": location,
                    "speed": f"{memory.Speed if hasattr(memory, 'Speed') and memory.Speed else 0} MHz"
                })
            
            # Atualiza o resultado
            if total_memory_gb > 0:
                result["total"] = f"{total_memory_gb:.2f} GB"
            
            if modules:
                result["slots_used"] = str(len(modules))
                result["modules"] = modules
        except Exception as e:
            print(f"Erro ao obter informações da memória RAM via WMI: {e}")
        
        return result
    
    def _get_ram_info_wmic(self):
        """Obtém informações da memória RAM usando WMIC."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações dos módulos de memória
            output = subprocess.check_output(
                ["wmic", "memorychip", "get", "capacity,manufacturer,banklabel,devicelocator,speed"],
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            lines = output.strip().split("\n")
            header = lines[0]
            
            # Determina os índices das colunas
            banklabel_index = header.lower().find("banklabel")
            capacity_index = header.lower().find("capacity")
            devicelocator_index = header.lower().find("devicelocator")
            manufacturer_index = header.lower().find("manufacturer")
            speed_index = header.lower().find("speed")
            
            # Ordena os índices para extrair corretamente
            indices = sorted([
                (banklabel_index, "banklabel"), 
                (capacity_index, "capacity"), 
                (devicelocator_index, "devicelocator"), 
                (manufacturer_index, "manufacturer"),
                (speed_index, "speed")
            ])
            
            total_memory_gb = 0
            modules = []
            
            for line in lines[1:]:
                if not line.strip():
                    continue
                
                module_data = {}
                last_index = 0
                for i, (index, key) in enumerate(indices):
                    if index < 0: continue
                    
                    # Calcula o fim da coluna atual
                    end_index = indices[i+1][0] if i + 1 < len(indices) and indices[i+1][0] >= 0 else len(line)
                    
                    # Extrai o valor
                    value = line[index:end_index].strip()
                    module_data[key] = value
                
                try:
                    # Converte a capacidade para GB
                    size_gb = round(int(module_data.get("capacity", 0)) / (1024**3), 2)
                    total_memory_gb += size_gb
                    
                    # Obtém o fabricante
                    manufacturer = module_data.get("manufacturer", "Não disponível")
                    if manufacturer.lower() in ["unknown", "not specified", "0000", "to be filled by o.e.m."]:
                        manufacturer = "Não identificado"
                    
                    # Obtém a localização
                    location = module_data.get("banklabel", "Não disponível")
                    if location == "Não disponível":
                        location = module_data.get("devicelocator", "Não disponível")
                    
                    # Adiciona o módulo à lista
                    modules.append({
                        "size": f"{size_gb} GB",
                        "manufacturer": manufacturer,
                        "location": location,
                        "speed": f"{module_data.get('speed', 0)} MHz"
                    })
                except (ValueError, IndexError):
                    continue
            
            # Atualiza o resultado
            if total_memory_gb > 0:
                result["total"] = f"{total_memory_gb:.2f} GB"
            
            if modules:
                result["slots_used"] = str(len(modules))
                result["modules"] = modules
        except Exception as e:
            print(f"Erro ao obter informações da memória RAM via WMIC: {e}")
        
        return result
    
    def _get_ram_info_psutil(self):
        """Obtém informações da memória RAM usando psutil."""
        result = {
            "total": "Não disponível",
            "slots_used": "Não disponível",
            "modules": []
        }
        
        if not psutil:
            return result
            
        try:
            # Obtém a memória total
            memory = psutil.virtual_memory()
            total_gb = round(memory.total / (1024**3), 2)
            
            result["total"] = f"{total_gb:.2f} GB"
            
            # Não é possível obter o número de slots ou informações dos módulos com psutil
            # Deixa como não disponível
            result["slots_used"] = "Não disponível"
            result["modules"] = []
        except Exception as e:
            print(f"Erro ao obter informações da memória RAM via psutil: {e}")
        
        return result
    
    def get_disk_info(self):
        """Obtém informações dos discos."""
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_disk_info_wmi,
            self._get_disk_info_wmic,
            self._get_disk_info_psutil
        ]
        
        all_disks = []
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                result = method()
                if result and len(result) > 0:
                    # Combina resultados se possível (evita duplicatas)
                    for disk in result:
                        # Verifica se um disco com modelo similar já existe
                        found = False
                        for existing_disk in all_disks:
                            if existing_disk["model"] == disk["model"]:
                                # Atualiza informações se forem mais completas
                                if disk["free_space"] != "Não disponível" and existing_disk["free_space"] == "Não disponível":
                                    existing_disk["free_space"] = disk["free_space"]
                                if disk["type"] != "Não disponível" and existing_disk["type"] == "Não disponível":
                                    existing_disk["type"] = disk["type"]
                                found = True
                                break
                        if not found:
                            all_disks.append(disk)
                    # Se já temos informações de todos os métodos, podemos parar
                    # (ou continuar para tentar obter mais detalhes)
            except Exception as e:
                print(f"Erro ao obter informações dos discos usando {method.__name__}: {e}")
        
        # Se nenhum método funcionar, retorna uma lista vazia
        return all_disks
    
    def _get_disk_info_wmi(self):
        """Obtém informações dos discos usando WMI."""
        result = []
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Obtém informações dos discos físicos
            physical_disks = self.wmi_client.Win32_DiskDrive()
            logical_disks = self.wmi_client.Win32_LogicalDisk(DriveType=3) # Tipo 3 = Disco local
            
            # Mapeia partições para discos físicos
            partition_map = {}
            for partition in self.wmi_client.Win32_DiskPartition():
                disk_index = partition.DiskIndex
                if disk_index not in partition_map:
                    partition_map[disk_index] = []
                partition_map[disk_index].append(partition.DeviceID)
                
            # Mapeia discos lógicos para partições
            logical_map = {}
            for logical in logical_disks:
                for partition_ref in logical.associators(wmi_result_class="Win32_LogicalDiskToPartition"):
                    partition_id = partition_ref.DeviceID
                    logical_map[partition_id] = {
                        "free_space": round(int(logical.FreeSpace) / (1024**3), 2) if hasattr(logical, "FreeSpace") and logical.FreeSpace else 0,
                        "total_space": round(int(logical.Size) / (1024**3), 2) if hasattr(logical, "Size") and logical.Size else 0
                    }
            
            for disk in physical_disks:
                # Calcula o tamanho em GB
                size_gb = round(int(disk.Size) / (1024**3), 2) if hasattr(disk, "Size") and disk.Size else 0
                
                # Determina o tipo de disco
                disk_type = "HDD"
                if hasattr(disk, "MediaType") and disk.MediaType:
                    media_type_str = str(disk.MediaType).lower()
                    if "ssd" in media_type_str or "solid state" in media_type_str:
                        disk_type = "SSD"
                    elif "nvme" in media_type_str:
                        disk_type = "NVMe"
                elif hasattr(disk, "Model") and disk.Model and ("ssd" in disk.Model.lower() or "nvme" in disk.Model.lower()):
                    # Tenta inferir pelo modelo se MediaType não for útil
                    if "nvme" in disk.Model.lower():
                        disk_type = "NVMe"
                    else:
                        disk_type = "SSD"
                
                # Obtém o modelo
                model = disk.Model.strip() if hasattr(disk, "Model") and disk.Model else "Não disponível"
                
                # Calcula o espaço livre total para este disco físico
                total_free_space_gb = 0
                disk_index = disk.Index
                if disk_index in partition_map:
                    for partition_id in partition_map[disk_index]:
                        if partition_id in logical_map:
                            total_free_space_gb += logical_map[partition_id]["free_space"]
                
                # Adiciona o disco à lista
                result.append({
                    "model": model,
                    "type": disk_type,
                    "size": f"{size_gb:.2f} GB",
                    "free_space": f"{total_free_space_gb:.2f} GB" if total_free_space_gb > 0 else "Não disponível"
                })
        except Exception as e:
            print(f"Erro ao obter informações dos discos via WMI: {e}")
        
        return result
    
    def _get_disk_info_wmic(self):
        """Obtém informações dos discos usando WMIC."""
        result = []
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações dos discos físicos
            model_output = subprocess.check_output(
                ["wmic", "diskdrive", "get", "model,size,mediatype"],
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            lines = model_output.strip().split("\n")
            header = lines[0]
            
            # Determina os índices das colunas
            mediatype_index = header.lower().find("mediatype")
            model_index = header.lower().find("model")
            size_index = header.lower().find("size")
            
            # Ordena os índices para extrair corretamente
            indices = sorted([
                (mediatype_index, "mediatype"), 
                (model_index, "model"), 
                (size_index, "size")
            ])
            
            for line in lines[1:]:
                if not line.strip():
                    continue
                
                disk_data = {}
                last_index = 0
                for i, (index, key) in enumerate(indices):
                    if index < 0: continue
                    
                    # Calcula o fim da coluna atual
                    end_index = indices[i+1][0] if i + 1 < len(indices) and indices[i+1][0] >= 0 else len(line)
                    
                    # Extrai o valor
                    value = line[index:end_index].strip()
                    disk_data[key] = value
                
                try:
                    # Converte o tamanho para GB
                    size_gb = round(int(disk_data.get("size", 0)) / (1024**3), 2)
                    
                    # Determina o tipo de disco
                    disk_type = "HDD"
                    media_type = disk_data.get("mediatype", "").lower()
                    model = disk_data.get("model", "").lower()
                    
                    if "ssd" in media_type or "solid state" in media_type or "ssd" in model:
                        disk_type = "SSD"
                    if "nvme" in media_type or "nvme" in model:
                         disk_type = "NVMe"
                    
                    # Adiciona o disco à lista
                    result.append({
                        "model": disk_data.get("model", "Não disponível"),
                        "type": disk_type,
                        "size": f"{size_gb:.2f} GB",
                        "free_space": "Não disponível" # WMIC não fornece espaço livre facilmente
                    })
                except (ValueError, IndexError):
                    continue
                    
            # Tenta obter espaço livre com outro comando WMIC
            try:
                logical_output = subprocess.check_output(
                    ["wmic", "logicaldisk", "where", "drivetype=3", "get", "deviceid,freespace,size"],
                    universal_newlines=True,
                    timeout=10
                )
                logical_lines = logical_output.strip().split("\n")
                logical_header = logical_lines[0]
                
                deviceid_index = logical_header.lower().find("deviceid")
                freespace_index = logical_header.lower().find("freespace")
                size_logical_index = logical_header.lower().find("size")
                
                logical_indices = sorted([
                    (deviceid_index, "deviceid"), 
                    (freespace_index, "freespace"), 
                    (size_logical_index, "size")
                ])
                
                total_free_space_gb = 0
                for logical_line in logical_lines[1:]:
                    if not logical_line.strip(): continue
                    
                    logical_data = {}
                    last_logical_index = 0
                    for i, (index, key) in enumerate(logical_indices):
                        if index < 0: continue
                        end_index = logical_indices[i+1][0] if i + 1 < len(logical_indices) and logical_indices[i+1][0] >= 0 else len(logical_line)
                        value = logical_line[index:end_index].strip()
                        logical_data[key] = value
                        
                    try:
                        free_space_gb = round(int(logical_data.get("freespace", 0)) / (1024**3), 2)
                        total_free_space_gb += free_space_gb
                    except (ValueError, IndexError):
                        continue
                
                # Atribui o espaço livre total ao primeiro disco (aproximação)
                if result and total_free_space_gb > 0:
                    result[0]["free_space"] = f"{total_free_space_gb:.2f} GB"
            except Exception as e:
                print(f"Erro ao obter espaço livre via WMIC: {e}")
        except Exception as e:
            print(f"Erro ao obter informações dos discos via WMIC: {e}")
        
        return result
    
    def _get_disk_info_psutil(self):
        """Obtém informações dos discos usando psutil."""
        result = []
        
        if not psutil:
            return result
            
        try:
            # Obtém informações das partições
            partitions = psutil.disk_partitions()
            
            for partition in partitions:
                try:
                    # Obtém informações de uso da partição
                    usage = psutil.disk_usage(partition.mountpoint)
                    
                    # Calcula tamanhos em GB
                    total_gb = round(usage.total / (1024**3), 2)
                    free_gb = round(usage.free / (1024**3), 2)
                    
                    # Adiciona o disco/partição à lista
                    # psutil não fornece modelo ou tipo facilmente
                    result.append({
                        "model": f"Partição {partition.device}",
                        "type": "Não disponível",
                        "size": f"{total_gb:.2f} GB",
                        "free_space": f"{free_gb:.2f} GB"
                    })
                except Exception as e:
                    print(f"Erro ao processar partição {partition.device}: {e}")
                    continue
        except Exception as e:
            print(f"Erro ao obter informações dos discos via psutil: {e}")
        
        return result
    
    def get_gpu_info(self):
        """Obtém informações da placa de vídeo."""
        result = []
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_gpu_info_wmi,
            self._get_gpu_info_wmic,
            self._get_gpu_info_powershell
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result and len(method_result) > 0:
                    # Combina resultados se possível (evita duplicatas)
                    for gpu in method_result:
                        found = False
                        for existing_gpu in result:
                            if existing_gpu["model"] == gpu["model"]:
                                found = True
                                break
                        if not found:
                            result.append(gpu)
            except Exception as e:
                print(f"Erro ao obter informações da GPU usando {method.__name__}: {e}")
        
        # Se nenhum método funcionar, retorna uma lista vazia
        return result
    
    def _get_gpu_info_wmi(self):
        """Obtém informações da placa de vídeo usando WMI."""
        result = []
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Obtém informações das placas de vídeo
            for gpu in self.wmi_client.Win32_VideoController():
                model = gpu.Name.strip() if hasattr(gpu, "Name") and gpu.Name else "Não disponível"
                
                # Adiciona a GPU à lista
                result.append({
                    "model": model
                })
        except Exception as e:
            print(f"Erro ao obter informações da GPU via WMI: {e}")
        
        return result
    
    def _get_gpu_info_wmic(self):
        """Obtém informações da placa de vídeo usando WMIC."""
        result = []
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações das placas de vídeo
            output = subprocess.check_output(
                ["wmic", "path", "win32_videocontroller", "get", "name"],
                universal_newlines=True,
                timeout=10
            )
            
            # Processa a saída
            lines = output.strip().split("\n")
            for line in lines[1:]:
                model = line.strip()
                if model:
                    result.append({
                        "model": model
                    })
        except Exception as e:
            print(f"Erro ao obter informações da GPU via WMIC: {e}")
        
        return result
        
    def _get_gpu_info_powershell(self):
        """Obtém informações da placa de vídeo usando PowerShell."""
        result = []
        
        if not self.is_windows:
            return result
        
        try:
            # Comando PowerShell para obter informações da GPU
            ps_command = "Get-CimInstance -ClassName Win32_VideoController | Select-Object -Property Name | ConvertTo-Json"
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode == 0 and stdout.strip():
                try:
                    data = json.loads(stdout)
                    if not isinstance(data, list):
                        data = [data]
                    
                    for gpu in data:
                        model = gpu.get("Name", "").strip()
                        if model:
                            result.append({
                                "model": model
                            })
                except json.JSONDecodeError:
                    print("Erro ao decodificar JSON da saída do PowerShell para GPU")
        except Exception as e:
            print(f"Erro ao obter informações da GPU via PowerShell: {e}")
        
        return result

    def get_display_info(self):
        """Obtém informações do display."""
        result = {
            "resolution": "Não disponível"
        }
        
        if not screeninfo:
            return result
            
        try:
            # Obtém informações do monitor principal
            primary_monitor = None
            for monitor in screeninfo.get_monitors():
                if monitor.is_primary:
                    primary_monitor = monitor
                    break
            
            # Se não encontrar o primário, usa o primeiro monitor
            if not primary_monitor and screeninfo.get_monitors():
                primary_monitor = screeninfo.get_monitors()[0]
            
            if primary_monitor:
                result["resolution"] = f"{primary_monitor.width}x{primary_monitor.height}"
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
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_tpm_info_wmi,
            self._get_tpm_info_powershell,
            self._get_tpm_info_tpmtool
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores foram preenchidos, interrompe a busca
                    if all(value != "Não disponível" for value in result.values()):
                        break
            except Exception as e:
                print(f"Erro ao obter informações do TPM usando {method.__name__}: {e}")
        
        return result
    
    def _get_tpm_info_wmi(self):
        """Obtém informações do TPM usando WMI."""
        result = {
            "version": "Não disponível",
            "status": "Não disponível",
            "manufacturer": "Não disponível"
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Tenta acessar o namespace correto
            tpm_info = self.wmi_client.Win32_Tpm()
            if tpm_info:
                tpm = tpm_info[0]
                # Versão
                if hasattr(tpm, "SpecVersion") and tpm.SpecVersion:
                    result["version"] = tpm.SpecVersion.strip()
                
                # Status
                if hasattr(tpm, "IsEnabled_InitialValue") and tpm.IsEnabled_InitialValue is not None:
                    result["status"] = "Habilitado" if tpm.IsEnabled_InitialValue else "Desabilitado"
                elif hasattr(tpm, "IsActivated_InitialValue") and tpm.IsActivated_InitialValue is not None:
                     result["status"] = "Ativo" if tpm.IsActivated_InitialValue else "Inativo"
                
                # Fabricante
                if hasattr(tpm, "ManufacturerIdTxt") and tpm.ManufacturerIdTxt:
                    result["manufacturer"] = tpm.ManufacturerIdTxt.strip()
                elif hasattr(tpm, "ManufacturerVersion") and tpm.ManufacturerVersion:
                    # Tenta extrair do ManufacturerVersion se ID não estiver disponível
                    result["manufacturer"] = tpm.ManufacturerVersion.split(",")[0].strip()
        except Exception as e:
            print(f"Erro ao obter informações do TPM via WMI: {e}")
        
        return result
    
    def _get_tpm_info_powershell(self):
        """Obtém informações do TPM usando PowerShell."""
        result = {
            "version": "Não disponível",
            "status": "Não disponível",
            "manufacturer": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Comando PowerShell para obter informações do TPM
            ps_command = "Get-Tpm | Select-Object -Property TpmPresent,TpmReady,TpmEnabled,TpmActivated,ManufacturerId,ManufacturerVersion,ManufacturerIdTxt | ConvertTo-Json"
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode == 0 and stdout.strip():
                try:
                    data = json.loads(stdout)
                    
                    # Versão
                    if "ManufacturerVersion" in data and data["ManufacturerVersion"]:
                        result["version"] = data["ManufacturerVersion"]
                    
                    # Status
                    if "TpmEnabled" in data:
                        result["status"] = "Habilitado" if data["TpmEnabled"] else "Desabilitado"
                    elif "TpmReady" in data:
                         result["status"] = "Pronto" if data["TpmReady"] else "Não pronto"
                    
                    # Fabricante
                    if "ManufacturerIdTxt" in data and data["ManufacturerIdTxt"]:
                        result["manufacturer"] = data["ManufacturerIdTxt"]
                    elif "ManufacturerId" in data and data["ManufacturerId"]:
                        result["manufacturer"] = str(data["ManufacturerId"]) # Converte ID numérico para string
                except json.JSONDecodeError:
                    # Tenta processar a saída como texto se não for JSON válido
                    lines = stdout.strip().split("\n")
                    for line in lines:
                        if "ManufacturerVersion" in line and ":" in line:
                            result["version"] = line.split(":", 1)[1].strip()
                        elif "TpmEnabled" in line and ":" in line:
                            value = line.split(":", 1)[1].strip().lower()
                            result["status"] = "Habilitado" if value == "true" else "Desabilitado"
                        elif "ManufacturerIdTxt" in line and ":" in line:
                            result["manufacturer"] = line.split(":", 1)[1].strip()
                        elif "ManufacturerId" in line and ":" in line:
                            result["manufacturer"] = line.split(":", 1)[1].strip()
        except Exception as e:
            print(f"Erro ao obter informações do TPM via PowerShell: {e}")
        
        return result
    
    def _get_tpm_info_tpmtool(self):
        """Obtém informações do TPM usando o TPM Tool."""
        result = {
            "version": "Não disponível",
            "status": "Não disponível",
            "manufacturer": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando TPM Tool
            output = subprocess.check_output(
                ["tpmtool", "getinfo"],
                universal_newlines=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Processa a saída
            for line in output.strip().split("\n"):
                if "Spec Version:" in line:
                    result["version"] = line.split(":", 1)[1].strip()
                elif "TPM Enabled:" in line:
                    value = line.split(":", 1)[1].strip().lower()
                    result["status"] = "Habilitado" if value == "yes" or value == "true" else "Desabilitado"
                elif "Manufacturer:" in line:
                    result["manufacturer"] = line.split(":", 1)[1].strip()
        except Exception as e:
            print(f"Erro ao obter informações do TPM via TPM Tool: {e}")
        
        return result
    
    def get_bluetooth_info(self):
        """Obtém informações do Bluetooth."""
        result = {
            "device_name": "Não disponível",
            "device_status": "Não disponível"
        }
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_bluetooth_info_wmi,
            self._get_bluetooth_info_wmic,
            self._get_bluetooth_info_powershell
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores foram preenchidos, interrompe a busca
                    if all(value != "Não disponível" for value in result.values()):
                        break
            except Exception as e:
                print(f"Erro ao obter informações do Bluetooth usando {method.__name__}: {e}")
        
        return result
    
    def _get_bluetooth_info_wmi(self):
        """Obtém informações do Bluetooth usando WMI."""
        result = {
            "device_name": "Não disponível",
            "device_status": "Não disponível"
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Procura por dispositivos Bluetooth
            for device in self.wmi_client.Win32_PnPEntity():
                if hasattr(device, "Name") and hasattr(device, "Status") and hasattr(device, "Description"):
                    name = device.Name if device.Name else ""
                    description = device.Description if device.Description else ""
                    
                    # Verifica se é um dispositivo Bluetooth
                    if "bluetooth" in name.lower() or "bluetooth" in description.lower():
                        result["device_name"] = name
                        result["device_status"] = device.Status if device.Status else "Desconhecido"
                        break
        except Exception as e:
            print(f"Erro ao obter informações do Bluetooth via WMI: {e}")
        
        return result
    
    def _get_bluetooth_info_wmic(self):
        """Obtém informações do Bluetooth usando WMIC."""
        result = {
            "device_name": "Não disponível",
            "device_status": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando para obter informações dos dispositivos PnP
            output = subprocess.check_output(
                ["wmic", "path", "Win32_PnPEntity", "where", "Name like '%Bluetooth%'", "get", "Name,Status"],
                universal_newlines=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Processa a saída
            lines = output.strip().split("\n")
            if len(lines) > 1:  # Pelo menos uma linha além do cabeçalho
                # Determina os índices das colunas
                header = lines[0]
                name_index = header.lower().find("name")
                status_index = header.lower().find("status")
                
                for line in lines[1:]:
                    if not line.strip():
                        continue
                    
                    # Extrai o nome e o status
                    if name_index >= 0 and status_index >= 0 and len(line) > max(name_index, status_index):
                        if name_index < status_index:
                            name = line[name_index:status_index].strip()
                            status = line[status_index:].strip()
                        else:
                            status = line[status_index:name_index].strip()
                            name = line[name_index:].strip()
                        
                        if name:
                            result["device_name"] = name
                            result["device_status"] = status if status else "Desconhecido"
                            break
        except Exception as e:
            print(f"Erro ao obter informações do Bluetooth via WMIC: {e}")
        
        return result
    
    def _get_bluetooth_info_powershell(self):
        """Obtém informações do Bluetooth usando PowerShell."""
        result = {
            "device_name": "Não disponível",
            "device_status": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Comando PowerShell para obter informações do Bluetooth
            ps_command = "Get-PnpDevice | Where-Object {$_.Name -like '*Bluetooth*'} | Select-Object -First 1 -Property Name,Status | ConvertTo-Json"
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode == 0 and stdout.strip():
                try:
                    data = json.loads(stdout)
                    
                    if "Name" in data and data["Name"]:
                        result["device_name"] = data["Name"]
                    
                    if "Status" in data and data["Status"]:
                        status = data["Status"]
                        if status == "OK":
                            result["device_status"] = "Ativo"
                        else:
                            result["device_status"] = status
                except json.JSONDecodeError:
                    # Tenta processar a saída como texto se não for JSON válido
                    lines = stdout.strip().split("\n")
                    for line in lines:
                        if "Name" in line and ":" in line:
                            result["device_name"] = line.split(":", 1)[1].strip()
                        elif "Status" in line and ":" in line:
                            status = line.split(":", 1)[1].strip()
                            if status == "OK":
                                result["device_status"] = "Ativo"
                            else:
                                result["device_status"] = status
        except Exception as e:
            print(f"Erro ao obter informações do Bluetooth via PowerShell: {e}")
        
        return result

    def get_wifi_info(self):
        """Obtém informações do Wi-Fi."""
        result = {
            "adapter_name": "Não disponível",
            "adapter_status": "Não disponível",
            "connected_ssid": "Não disponível"
        }
        
        # Lista de métodos para tentar obter as informações
        methods = [
            self._get_wifi_info_wmi,
            self._get_wifi_info_netsh,
            self._get_wifi_info_powershell
        ]
        
        # Tenta cada método até obter sucesso
        for method in methods:
            try:
                method_result = method()
                if method_result:
                    # Atualiza apenas os valores que foram obtidos com sucesso
                    for key, value in method_result.items():
                        if value and value != "Não disponível":
                            result[key] = value
                    
                    # Se todos os valores foram preenchidos, interrompe a busca
                    if all(value != "Não disponível" for value in result.values()):
                        break
            except Exception as e:
                print(f"Erro ao obter informações do Wi-Fi usando {method.__name__}: {e}")
        
        return result
    
    def _get_wifi_info_wmi(self):
        """Obtém informações do Wi-Fi usando WMI."""
        result = {
            "adapter_name": "Não disponível",
            "adapter_status": "Não disponível",
            "connected_ssid": "Não disponível"
        }
        
        if not self.wmi_client or not self.is_windows:
            return result
        
        try:
            # Procura por adaptadores de rede sem fio
            for adapter in self.wmi_client.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if hasattr(adapter, "Description") and "wireless" in adapter.Description.lower():
                    result["adapter_name"] = adapter.Description.strip()
                    
                    # Obtém o status do adaptador físico correspondente
                    physical_adapter = self.wmi_client.Win32_NetworkAdapter(Index=adapter.Index)
                    if physical_adapter and hasattr(physical_adapter[0], "NetConnectionStatus"):
                        status_map = {
                            0: "Desconectado", 1: "Conectando", 2: "Conectado",
                            3: "Desconectando", 4: "Hardware não presente",
                            5: "Hardware desabilitado", 6: "Erro de hardware",
                            7: "Mídia desconectada", 8: "Autenticando",
                            9: "Autenticação bem-sucedida", 10: "Falha na autenticação",
                            11: "Endereço inválido", 12: "Credenciais necessárias"
                        }
                        result["adapter_status"] = status_map.get(physical_adapter[0].NetConnectionStatus, "Desconhecido")
                    
                    # Tenta obter o SSID conectado (pode falhar sem privilégios)
                    try:
                        ssid_info = self.wmi_client.query(
                            f"SELECT SSID FROM MSNdis_80211_ServiceSetIdentifier WHERE active=true AND InstanceName='{physical_adapter[0].GUID}'"
                        )
                        if ssid_info:
                            result["connected_ssid"] = ssid_info[0].SSID
                    except Exception:
                        pass # Ignora erro ao obter SSID
                        
                    break # Pega o primeiro adaptador wireless encontrado
        except Exception as e:
            print(f"Erro ao obter informações do Wi-Fi via WMI: {e}")
        
        return result
    
    def _get_wifi_info_netsh(self):
        """Obtém informações do Wi-Fi usando netsh."""
        result = {
            "adapter_name": "Não disponível",
            "adapter_status": "Não disponível",
            "connected_ssid": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Executa o comando netsh para obter informações das interfaces WLAN
            output = subprocess.check_output(
                ["netsh", "wlan", "show", "interfaces"],
                universal_newlines=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Processa a saída
            adapter_name = "Não disponível"
            status = "Não disponível"
            ssid = "Não disponível"
            
            for line in output.strip().split("\n"):
                line = line.strip()
                if line.startswith("Name"):
                    adapter_name = line.split(":", 1)[1].strip()
                elif line.startswith("State"):
                    status_raw = line.split(":", 1)[1].strip()
                    if status_raw == "connected":
                        status = "Conectado"
                    elif status_raw == "disconnected":
                        status = "Desconectado"
                    else:
                        status = status_raw
                elif line.startswith("SSID"):
                    ssid = line.split(":", 1)[1].strip()
            
            # Atualiza o resultado se encontrou um adaptador
            if adapter_name != "Não disponível":
                result["adapter_name"] = adapter_name
                result["adapter_status"] = status
                result["connected_ssid"] = ssid if status == "Conectado" else "Não conectado"
        except Exception as e:
            print(f"Erro ao obter informações do Wi-Fi via netsh: {e}")
        
        return result
        
    def _get_wifi_info_powershell(self):
        """Obtém informações do Wi-Fi usando PowerShell."""
        result = {
            "adapter_name": "Não disponível",
            "adapter_status": "Não disponível",
            "connected_ssid": "Não disponível"
        }
        
        if not self.is_windows:
            return result
        
        try:
            # Comando PowerShell para obter informações do Wi-Fi
            ps_command = "Get-NetAdapter -Name *Wi-Fi* | Select-Object -First 1 -Property Name,Status,InterfaceDescription | ConvertTo-Json; " \
                         "(Get-NetConnectionProfile -InterfaceAlias *Wi-Fi*).Name"
            
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            stdout, stderr = process.communicate(timeout=10)
            
            if process.returncode == 0 and stdout.strip():
                # Separa a saída do adaptador e do SSID
                parts = stdout.strip().split("\n")
                adapter_json = parts[0]
                ssid = parts[1] if len(parts) > 1 else "Não disponível"
                
                try:
                    adapter_data = json.loads(adapter_json)
                    
                    if "InterfaceDescription" in adapter_data and adapter_data["InterfaceDescription"]:
                        result["adapter_name"] = adapter_data["InterfaceDescription"]
                    elif "Name" in adapter_data and adapter_data["Name"]:
                         result["adapter_name"] = adapter_data["Name"]
                    
                    if "Status" in adapter_data and adapter_data["Status"]:
                        result["adapter_status"] = adapter_data["Status"]
                        
                    if result["adapter_status"] == "Up" or result["adapter_status"] == "Connected":
                         result["connected_ssid"] = ssid if ssid else "Não disponível"
                    else:
                        result["connected_ssid"] = "Não conectado"
                except json.JSONDecodeError:
                    print("Erro ao decodificar JSON da saída do PowerShell para Wi-Fi")
        except Exception as e:
            print(f"Erro ao obter informações do Wi-Fi via PowerShell: {e}")
        
        return result
