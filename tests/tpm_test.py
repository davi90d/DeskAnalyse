"""
Módulo para teste de TPM (Trusted Platform Module).
Implementação real que interage diretamente com o hardware.
"""

import os
import time
import subprocess
import re
import platform
import ctypes
from datetime import datetime

class TPMTest:
    """Classe para teste de TPM (Trusted Platform Module)."""
    
    def __init__(self):
        """Inicializa o teste de TPM."""
        self.result = {
            'success': False,
            'execution_time': 0,
            'version': 'Não disponível',
            'status': 'Não disponível',
            'manufacturer': 'Não disponível',
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
        """Executa o teste de TPM."""
        start_time = time.time()
        
        try:
            # Método 1: Usando PowerShell Get-Tpm
            if self._check_tpm_powershell():
                self.result['success'] = True
            # Método 2: Usando WMI
            elif self._check_tpm_wmi():
                self.result['success'] = True
            # Método 3: Usando tpm.msc e capturando saída
            elif self._check_tpm_msc():
                self.result['success'] = True
            else:
                self.result['error'] = "Não foi possível obter informações do TPM."
                self.result['success'] = False
        except Exception as e:
            self.result['error'] = f"Erro ao executar teste de TPM: {e}"
            self.result['success'] = False
        
        self.result['execution_time'] = time.time() - start_time
        return self.result['success']
    
    def _check_tpm_powershell(self):
        """Verifica o TPM usando PowerShell."""
        try:
            # Comando PowerShell para obter informações do TPM
            ps_command = "Get-Tpm | ConvertTo-Json"
            process = subprocess.Popen(
                ["powershell", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            stdout, stderr = process.communicate()
            
            if process.returncode == 0 and stdout.strip():
                import json
                tpm_info = json.loads(stdout)
                
                # Extrai informações
                self.result['status'] = 'Ativo' if tpm_info.get('TpmPresent', False) else 'Inativo'
                self.result['version'] = tpm_info.get('TpmVersion', 'Não disponível')
                
                # Tenta obter mais informações detalhadas
                self._get_tpm_manufacturer_info()
                
                return True
            return False
        except Exception as e:
            print(f"Erro ao verificar TPM via PowerShell: {e}")
            return False
    
    def _check_tpm_wmi(self):
        """Verifica o TPM usando WMI."""
        try:
            import wmi
            wmi_client = wmi.WMI()
            
            # Tenta acessar Win32_Tpm
            tpm_devices = wmi_client.Win32_Tpm()
            
            if tpm_devices:
                for tpm in tpm_devices:
                    self.result['status'] = 'Ativo' if tpm.IsEnabled_InitialValue else 'Inativo'
                    self.result['version'] = f"{tpm.SpecVersion}" if hasattr(tpm, 'SpecVersion') else 'Não disponível'
                    self.result['manufacturer'] = tpm.ManufacturerIdTxt if hasattr(tpm, 'ManufacturerIdTxt') else 'Não disponível'
                return True
            
            # Tenta acessar Win32_TPM (nome alternativo)
            tpm_devices = wmi_client.Win32_TPM()
            
            if tpm_devices:
                for tpm in tpm_devices:
                    self.result['status'] = 'Ativo' if tpm.IsEnabled_InitialValue else 'Inativo'
                    self.result['version'] = f"{tpm.SpecVersion}" if hasattr(tpm, 'SpecVersion') else 'Não disponível'
                    self.result['manufacturer'] = tpm.ManufacturerIdTxt if hasattr(tpm, 'ManufacturerIdTxt') else 'Não disponível'
                return True
            
            return False
        except Exception as e:
            print(f"Erro ao verificar TPM via WMI: {e}")
            return False
    
    def _check_tpm_msc(self):
        """Verifica o TPM usando tpm.msc e capturando saída."""
        try:
            # Cria um arquivo temporário para a saída
            temp_file = os.path.join(os.environ.get('TEMP', '.'), 'tpm_info.txt')
            
            # Executa o comando para obter informações do TPM
            subprocess.call(["tpm.msc"], shell=True)
            
            # Aguarda um tempo para o usuário visualizar e fechar
            time.sleep(2)
            
            # Tenta obter informações do TPM usando outro método
            ps_command = """
            $tpmWmi = Get-WmiObject -Namespace "root\\CIMV2\\Security\\MicrosoftTpm" -Class "Win32_Tpm"
            if ($tpmWmi) {
                $tpmInfo = @{
                    IsPresent = $true
                    IsEnabled = $tpmWmi.IsEnabled_InitialValue
                    IsActivated = $tpmWmi.IsActivated_InitialValue
                    SpecVersion = $tpmWmi.SpecVersion
                    ManufacturerId = $tpmWmi.ManufacturerId
                    ManufacturerVersion = $tpmWmi.ManufacturerVersion
                }
                $tpmInfo | ConvertTo-Json
            } else {
                @{ IsPresent = $false } | ConvertTo-Json
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
                    tpm_info = json.loads(stdout)
                    
                    if tpm_info.get('IsPresent', False):
                        self.result['status'] = 'Ativo' if tpm_info.get('IsEnabled', False) else 'Inativo'
                        self.result['version'] = tpm_info.get('SpecVersion', 'Não disponível')
                        
                        # Tenta interpretar o ID do fabricante
                        manufacturer_id = tpm_info.get('ManufacturerId')
                        if manufacturer_id:
                            self.result['manufacturer'] = self._get_manufacturer_name(manufacturer_id)
                        
                        return True
                except json.JSONDecodeError:
                    pass
            
            return False
        except Exception as e:
            print(f"Erro ao verificar TPM via tpm.msc: {e}")
            return False
    
    def _get_tpm_manufacturer_info(self):
        """Obtém informações do fabricante do TPM."""
        try:
            # Comando PowerShell para obter informações detalhadas do TPM
            ps_command = """
            $tpmWmi = Get-WmiObject -Namespace "root\\CIMV2\\Security\\MicrosoftTpm" -Class "Win32_Tpm"
            if ($tpmWmi -and $tpmWmi.ManufacturerId) {
                $tpmWmi.ManufacturerId
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
                manufacturer_id = stdout.strip()
                self.result['manufacturer'] = self._get_manufacturer_name(manufacturer_id)
        except Exception as e:
            print(f"Erro ao obter informações do fabricante do TPM: {e}")
    
    def _get_manufacturer_name(self, manufacturer_id):
        """Converte o ID do fabricante para o nome."""
        # Dicionário de IDs de fabricantes conhecidos
        manufacturers = {
            "0x414D4400": "AMD",
            "0x41544D4C": "Atmel",
            "0x4252434D": "Broadcom",
            "0x49424d00": "IBM",
            "0x49465800": "Infineon",
            "0x494E5443": "Intel",
            "0x4C454E00": "Lenovo",
            "0x4E534D20": "National Semiconductor",
            "0x4E545A00": "Nationz",
            "0x4E544300": "Nuvoton Technology",
            "0x51434F4D": "Qualcomm",
            "0x534D5343": "SMSC",
            "0x53544D20": "ST Microelectronics",
            "0x534D534E": "Samsung",
            "0x534E5300": "Sinosun",
            "0x54584E00": "Texas Instruments",
            "0x57454300": "Winbond",
            "0x524F4343": "Fuzhou Rockchip"
        }
        
        # Tenta converter o ID para hexadecimal se não estiver nesse formato
        if not manufacturer_id.startswith("0x"):
            try:
                # Converte para inteiro e depois para hexadecimal
                manufacturer_id = f"0x{int(manufacturer_id):08X}"
            except:
                pass
        
        return manufacturers.get(manufacturer_id, f"Desconhecido ({manufacturer_id})")
    
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
            f"Teste de TPM: {status}\n"
            f"Tempo de execução: {self.result['execution_time']:.2f} segundos\n"
            f"Versão: {self.result['version']}\n"
            f"Status: {self.result['status']}\n"
            f"Fabricante: {self.result['manufacturer']}"
        )


# Função para teste rápido
if __name__ == "__main__":
    tpm_test = TPMTest()
    
    if tpm_test.initialize():
        print("Executando teste de TPM...")
        tpm_test.execute()
        print(tpm_test.get_formatted_result())
    else:
        print(f"Não foi possível inicializar o teste: {tpm_test.result['error']}")
