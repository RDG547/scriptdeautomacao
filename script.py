import subprocess
import platform
import os
import sys
import shutil
import logging
import time
import requests
import winsound
import winshell
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError
from elevate import elevate

# Configuração de logging
def configurar_logging():
    """Configura o sistema de logging."""
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler('meu_script.log', mode='w')
    file_handler.setLevel(logging.DEBUG)  # Logs detalhados no arquivo
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)  # Logs simplificados no terminal

    formatter = logging.Formatter('%(asctime)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger

logger = configurar_logging()

# Constantes
GERENCIADOR_DE_PACOTES_CHOCOLATEY = "choco"
URL_DO_SCRIPT_DE_INSTALACAO_DO_CHOCOLATEY = "https://chocolatey.org/install.ps1"
SERVIDOR_KMS = "kms8.msguides.com"
CAMINHO_CSCRIPT = r"C:\Windows\System32\cscript.exe"
CAMINHO_SLMGR = r"C:\Windows\System32\slmgr.vbs"
CAMINHO_CHOCOLATEY = shutil.which(GERENCIADOR_DE_PACOTES_CHOCOLATEY) or r"C:\ProgramData\chocolatey\bin\choco.exe"
COMANDO_POWERSHELL = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command"]

# Verificação de software instalado por executável
def verificar_instalado_por_executavel(caminho_executavel):
    """Verifica se um executável está presente no sistema."""
    logger.debug(f"Verificando se o executável existe: {caminho_executavel}")
    return os.path.exists(caminho_executavel)

def executar_e_logar(comando, mensagem_erro, validar_saida=None):
    """Executa um comando, registra logs detalhados e avalia a saída.

    Args:
        comando (list): O comando a ser executado.
        mensagem_erro (str): Mensagem a ser exibida em caso de erro.
        validar_saida (callable, opcional): Função para validar a saída do comando.

    Returns:
        subprocess.CompletedProcess: Resultado do comando executado.
    """
    try:
        logger.debug(f"Executando comando: {' '.join(comando)}")
        resultado = subprocess.run(
            comando, capture_output=True, text=True, stdin=subprocess.DEVNULL
        )

        logger.debug(f"Código de retorno: {resultado.returncode}")
        logger.debug(f"Saída padrão: {resultado.stdout.strip()}")
        logger.debug(f"Saída de erro: {resultado.stderr.strip()}")

        if resultado.returncode != 0:
            logger.error(f"{mensagem_erro}: {resultado.stderr.strip()} (código {resultado.returncode})")
        elif validar_saida and not validar_saida(resultado.stdout):
            logger.error(f"{mensagem_erro}: Validação de saída falhou.")
            return None

        return resultado
    except FileNotFoundError as fnf_error:
        logger.error(f"Arquivo não encontrado ao executar o comando: {comando[0]} - {fnf_error}")
    except PermissionError as perm_error:
        logger.error(f"Permissão negada ao executar o comando: {comando[0]} - {perm_error}")
    except Exception as e:
        logger.error(f"Erro inesperado ao executar o comando: {e}")
    return None

# Exemplo de validação de saída
def validar_saida_ativacao(stdout):
    """Valida se a ativação do Windows foi bem-sucedida pela saída."""
    return "Status da Licença: Licenciado" in stdout or "License Status: Licensed" in stdout

# Função para criar um atalho específico
def criar_atalho(nome, caminho_executavel, caminho_area_trabalho):
    caminho_atalho = os.path.join(caminho_area_trabalho, f"{nome}.lnk")
    
    logging.debug(f"Verificando existência do atalho: {caminho_atalho}")
    
    # Verificar se o atalho já existe como um arquivo
    if os.path.isfile(caminho_atalho):
        return
    
    try:
        with winshell.Shortcut(caminho_atalho) as atalho:
            atalho.path = caminho_executavel
            atalho.working_directory = os.path.dirname(caminho_executavel)
            atalho.write()
    except Exception as e:
        logging.error(f"Erro ao criar atalho para {nome}: {e}")

# Ativação do Office
def ativar_office():
    """Ativa o Microsoft Office 2019-2021 diretamente via comandos."""
    print("Ativando pacote Office...")
    # Diretórios onde o ospp.vbs pode estar localizado
    possiveis_caminhos = [
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16"
    ]

    # Encontrar o diretório correto
    diretorio_office = next((caminho for caminho in possiveis_caminhos if os.path.exists(os.path.join(caminho, "ospp.vbs"))), None)

    if not diretorio_office:
        print("Erro: Não foi possível localizar o arquivo ospp.vbs. O Microsoft Office pode não estar instalado corretamente.")
        return

    os.chdir(diretorio_office)

    try:
        # Inserir licenças disponíveis
        licenca_dir = os.path.join(diretorio_office, "..", "root", "Licenses16")
        if os.path.exists(licenca_dir):
            for licenca in os.listdir(licenca_dir):
                if licenca.startswith("ProPlus2021VL_KMS") and licenca.endswith(".xrm-ms"):
                    subprocess.run(["cscript", "ospp.vbs", f"/inslic:{os.path.join(licenca_dir, licenca)}"], check=False, stdout=subprocess.DEVNULL)
        else:
            print("Diretório de licenças não encontrado.")

        # Configurar servidor KMS
        subprocess.run(["cscript", "ospp.vbs", "/setprt:1688"], check=True, stdout=subprocess.DEVNULL)
        subprocess.run(["cscript", "ospp.vbs", "/unpkey:6F7TH"], check=False, stdout=subprocess.DEVNULL)

        servidores_kms = ["e8.us.to", "e9.us.to"]
        ativado = False

        for servidor in servidores_kms:
            subprocess.run(["cscript", "ospp.vbs", f"/sethst:{servidor}"], check=True, stdout=subprocess.DEVNULL)

            # Inserir chave de produto
            chave_produto = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
            subprocess.run(["cscript", "ospp.vbs", f"/inpkey:{chave_produto}"], check=True, stdout=subprocess.DEVNULL)

            # Tentar ativar
            resultado = subprocess.run(["cscript", "ospp.vbs", "/act"], capture_output=True, text=True)

            if "successful" in resultado.stdout.lower():
                print(f"Ativação do pacote Office concluída!")
                ativado = True
                break
            else:
                print(f"Falha na ativação com o servidor {servidor}. Tentando outro servidor...")

        if not ativado:
            print("Falha na ativação! Todos os servidores de ativação foram tentados sem sucesso.")

    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o processo de ativação: {e}")

# Classe AtivadorWindows
class AtivadorWindows:
    """Classe responsável por verificar e ativar o Windows."""

    CHAVE_WINDOWS_PRO = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
    CHAVE_WINDOWS_HOME = "7KNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"

    @staticmethod
    def verificar_ativacao_do_windows():
        """Verifica se o Windows está ativado."""
        resultado = executar_comando([CAMINHO_CSCRIPT, "//Nologo", CAMINHO_SLMGR, "/dlv"], "Erro ao verificar ativação do Windows")

        if resultado and resultado.stdout:
            saida = resultado.stdout.replace('Æ', 'ç').replace('‡', 'ç').replace('ä', 'á').replace('ˆ', 'ê')
            if "Status da Licença: Licenciado" in saida or "License Status: Licensed" in saida:
                return True
        return False

    @staticmethod
    def escolher_chave_produto():
        """Escolhe a chave de produto com base na edição do Windows."""
        try:
            edicao_windows = platform.win32_edition()
            logging.info(f"Edição do Windows detectada: {edicao_windows}")

            if edicao_windows in ['Pro', 'Professional']:
                logging.info("Usando chave de produto para Windows Pro.")
                return AtivadorWindows.CHAVE_WINDOWS_PRO
            elif edicao_windows in ['Home', 'Core', 'CoreSingleLanguage']:
                logging.info("Usando chave de produto para Windows Home.")
                return AtivadorWindows.CHAVE_WINDOWS_HOME
            else:
                logging.error(f"Edição do Windows não suportada: {edicao_windows}")
                return None
        except Exception as e:
            logging.error(f"Erro ao determinar a edição do Windows: {e}")
            return None

    @staticmethod
    def ativar_windows():
        """Ativa o Windows usando o servidor KMS especificado."""

        chave_produto = AtivadorWindows.escolher_chave_produto()
        if not chave_produto:
            logging.error("Chave de produto não encontrada. A ativação foi abortada.")
            return

        executar_e_logar([CAMINHO_CSCRIPT, "//Nologo", CAMINHO_SLMGR, "/upk"], "Erro ao desinstalar chave de produto")
        executar_e_logar([CAMINHO_CSCRIPT, "//Nologo", CAMINHO_SLMGR, "/ipk", chave_produto], "Erro ao instalar chave de produto")
        executar_e_logar([CAMINHO_CSCRIPT, "//Nologo", CAMINHO_SLMGR, "/skms", SERVIDOR_KMS], "Erro ao configurar servidor KMS")
        executar_e_logar([CAMINHO_CSCRIPT, "//Nologo", CAMINHO_SLMGR, "/ato"], "Erro ao ativar Windows")

    @staticmethod
    def verificar_e_ativar_windows():
        """Verifica a ativação do Windows e tenta ativar se não estiver ativado."""
        logging.info("Verificando ativação do Windows...")

        if AtivadorWindows.verificar_ativacao_do_windows():
            logging.info("O Windows já está ativado.")
        else:
            logging.info("Ativando o Windows...")
            AtivadorWindows.ativar_windows()

            # Verificar novamente após tentar ativar
            if AtivadorWindows.verificar_ativacao_do_windows():
                logging.info("Ativação do Windows concluída.")
            else:
                logging.error("Erro ao ativar o Windows.")

# Dicionário de configuração de software
CONFIGURACAO_DE_SOFTWARE = {
    'Adobe Acrobat Reader': {
        'caminho_executavel': r'C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe',
        'argumentos_de_instalacao': ['choco', 'install', '-y', 'adobereader', '--ignore-pending-reboot', '--no-progress']
    },
    'Google Chrome': {
        'caminho_executavel': r'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'argumentos_de_instalacao': ['choco', 'install', '-y', 'googlechrome', '--params', '/NoDesktopShortcut', '--ignore-pending-reboot', '--no-progress']
    },
    'Mozilla Firefox': {
        'caminho_executavel': r'C:\\Program Files\\Mozilla Firefox\\firefox.exe',
        'argumentos_de_instalacao': ['choco', 'install', '-y', 'firefox', '--params', '/NoDesktopShortcut', '--ignore-pending-reboot', '--no-progress']
    },
    'WinRAR': {
        'caminho_executavel': r'C:\\Program Files\\WinRAR\\WinRAR.exe',
        'argumentos_de_instalacao': ['choco', 'install', '-y', 'winrar', '--ignore-pending-reboot', '--no-progress']
    },
    'Microsoft Office': {
        'caminho_executavel': r'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE',
        'argumentos_de_instalacao': [
            'choco', 'install', '-y', 'office2019proplus',
            '--params', '/S /ALLUSERS /LANG pt-br /NOAUTOACTIVATE /NOUPDATE',
            '--ignore-pending-reboot', '--no-progress'
        ]
    },
    'Word': {
        'caminho_executavel': r'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE',
    },
    'Excel': {
        'caminho_executavel': r'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
    },
    'PowerPoint': {
        'caminho_executavel': r'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE',
    },
}

# Variável para rastrear as atividades executadas
atividades_realizadas = []

# Função genérica para executar comandos
def executar_comando(comando, mensagem_de_erro):
    """Executa um comando e registra erro se houver."""
    try:
        env = os.environ.copy()
        env["PATH"] = f"{os.path.dirname(CAMINHO_CHOCOLATEY)};{env['PATH']}"

        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            stdin=subprocess.DEVNULL,
            env=env  # Passa o ambiente configurado
        )

        # Avaliar sucesso ou falha com base no código de retorno
        if resultado.returncode != 0:
            logging.error(f"{mensagem_de_erro}: Código {resultado.returncode}, Saída: {resultado.stdout or 'Nenhuma saída capturada.'}")
        return resultado
    except Exception as e:
        logging.error(f"{mensagem_de_erro}: {e}")
        return None

class InstaladorSoftware:
    """Classe responsável por instalar softwares via Chocolatey."""

    @staticmethod
    def verificar_instalacao_do_chocolatey():
        """Verifica se o Chocolatey está instalado e o instala se não estiver."""
        if shutil.which(GERENCIADOR_DE_PACOTES_CHOCOLATEY) and os.path.exists(CAMINHO_CHOCOLATEY):
            return True
        else:
            return InstaladorSoftware.instalar_chocolatey()

    @staticmethod
    def instalar_chocolatey():
        """Instala o Chocolatey usando um script de instalação remoto e atualiza o PATH."""
        comando_de_instalacao = [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-InputFormat", "None",
            "-ExecutionPolicy", "Bypass",
            "-Command", "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"
        ]
        resultado = subprocess.run(comando_de_instalacao, capture_output=True, text=True)
        if resultado.returncode == 0:
            # Atualiza o caminho para incluir o diretório do Chocolatey
            chocolatey_path = r"C:\ProgramData\chocolatey\bin"
            os.environ["PATH"] = f"{chocolatey_path};{os.environ['PATH']}"

            # Atualiza o valor global para CAMINHO_CHOCOLATEY
            global CAMINHO_CHOCOLATEY
            CAMINHO_CHOCOLATEY = shutil.which("choco") or r"C:\ProgramData\chocolatey\bin\choco.exe"

            return True
        else:
            logging.error(f"Erro ao instalar Chocolatey: {resultado.stderr}")
            return False

    @staticmethod
    def instalar_software(nome_do_software):
        """Instala o software especificado via Chocolatey e cria atalho."""
        logging.info(f"Instalando {nome_do_software}...")
        config = CONFIGURACAO_DE_SOFTWARE.get(nome_do_software)
        if not config:
            logging.error(f"Configuração para {nome_do_software} não encontrada.")
            return False

        # Tratar Word, Excel e PowerPoint como parte do Microsoft Office
        if nome_do_software in ['Word', 'Excel', 'PowerPoint']:
            return True

        # Verificar se o software já está instalado pelo executável
        if verificar_instalado_por_executavel(config['caminho_executavel']):
            logging.info(f"{nome_do_software} já está instalado.")
            if nome_do_software != "WinRAR":
                criar_atalho(nome_do_software, config['caminho_executavel'], winshell.desktop())
            else:
                # Remove atalho do WinRAR se já existir
                atalho_winrar = os.path.join(winshell.desktop(), "WinRAR.lnk")
                if os.path.exists(atalho_winrar):
                    os.remove(atalho_winrar)
            return True

        # Tentar instalar o software
        try:
            resultado = subprocess.run(config['argumentos_de_instalacao'], capture_output=True, text=True)
            if resultado.returncode == 0:
                logging.info(f"{nome_do_software} instalado com sucesso.")

                # Remover atalho padrão genérico do Microsoft Office
                if nome_do_software == "Microsoft Office":
                    atalho_office = os.path.join(winshell.desktop(), "Microsoft Office.lnk")
                    if os.path.exists(atalho_office):
                        os.remove(atalho_office)
                    # Criar atalhos específicos para Word, Excel e PowerPoint
                    criar_atalho("Word", CONFIGURACAO_DE_SOFTWARE["Word"]['caminho_executavel'], winshell.desktop())
                    criar_atalho("Excel", CONFIGURACAO_DE_SOFTWARE["Excel"]['caminho_executavel'], winshell.desktop())
                    criar_atalho("PowerPoint", CONFIGURACAO_DE_SOFTWARE["PowerPoint"]['caminho_executavel'], winshell.desktop())
                elif nome_do_software == "Google Chrome":
                    # Remover atalho padrão criado pelo instalador do Chrome
                    atalho_padrao_chrome = os.path.join(winshell.desktop(), "Google Chrome.lnk")
                    if os.path.exists(atalho_padrao_chrome):
                        os.remove(atalho_padrao_chrome)
                elif nome_do_software != "WinRAR":
                    # Criar atalho para outros softwares
                    criar_atalho(nome_do_software, config['caminho_executavel'], winshell.desktop())
        
                return True
            else:
                logging.error(f"Erro ao instalar {nome_do_software}: {resultado.stderr}")
                return False
        except Exception as e:
            logging.error(f"Erro inesperado ao instalar {nome_do_software}: {e}")
            return False

    @staticmethod
    def instalar_google_chrome_alternativo():
        """Tenta instalar o Google Chrome por fontes alternativas."""
        logging.info("Tentando instalar o Google Chrome de uma fonte alternativa...")

        url_instalador = "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"
        caminho_instalador = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "GoogleChromeInstaller.msi")

        try:
            # Baixar o instalador
            logging.info("Baixando o instalador do Google Chrome...")
            response = requests.get(url_instalador, stream=True)
            if response.status_code == 200:
                with open(caminho_instalador, "wb") as file:
                    file.write(response.content)
                logging.info("Instalador do Google Chrome baixado com sucesso.")

                # Executar o instalador
                logging.info("Executando o instalador do Google Chrome...")
                resultado = subprocess.run(["msiexec.exe", "/i", caminho_instalador, "/quiet", "/norestart"], capture_output=True, text=True)
                if resultado.returncode == 0:
                    logging.info("Google Chrome instalado com sucesso pela fonte alternativa.")
                    return True
                else:
                    logging.error(f"Erro ao instalar o Google Chrome pela fonte alternativa: {resultado.stderr}")
                    return False
            else:
                logging.error(f"Erro ao baixar o instalador do Google Chrome. Status HTTP: {response.status_code}")
                return False
        except Exception as e:
            logging.error(f"Erro durante a instalação alternativa do Google Chrome: {e}")
            return False

    @staticmethod
    def remover_chocolatey():
        """Remove completamente o Chocolatey e seus arquivos residuais."""

        # Desinstalar o Chocolatey via comando
        executar_comando(['choco', 'uninstall', 'chocolatey', '-y'], "Erro ao desinstalar o Chocolatey")

        # Remover arquivos residuais
        pasta_choco = r"C:\ProgramData\chocolatey"
        if os.path.exists(pasta_choco):
            shutil.rmtree(pasta_choco)

        # Remover o registro do Chocolatey
        chave_registro_choco = r"HKLM\Software\Chocolatey"
        try:
            # Verifica se a chave do registro existe antes de tentar removê-la
            comando_verificar_registro = COMANDO_POWERSHELL + [f"Test-Path '{chave_registro_choco}'"]
            resultado = executar_comando(comando_verificar_registro, "Erro ao verificar o registro do Chocolatey")

            if resultado and resultado.stdout.strip().lower() == 'true':
                # Se o caminho do registro existe, tentar remover
                comando_remover_registro = COMANDO_POWERSHELL + [f"Remove-Item -Path '{chave_registro_choco}' -Recurse -Force"]
                executar_comando(comando_remover_registro, "Erro ao remover o registro do Chocolatey")
        except Exception as e:
            logging.error(f"Erro ao remover o registro do Chocolatey: {e}")

class ConfiguradorWindows:
    """Classe responsável por configuração de aparência e desempenho do Windows usando pyautogui."""

    @staticmethod
    def verificar_estado_servico_windows_update():
        """Verifica as propriedades do serviço Windows Update para obter detalhes."""
        logging.info("Verificando propriedades do serviço Windows Update...")

        # Comando para obter o estado detalhado do serviço
        comando_verificar_detalhes = COMANDO_POWERSHELL + [
            "Get-Service -Name wuauserv | Format-List -Property *"
        ]
        resultado_detalhes = executar_comando(comando_verificar_detalhes, "Erro ao verificar as propriedades do serviço Windows Update")

        if resultado_detalhes and resultado_detalhes.stdout:
            logging.info(f"Propriedades do serviço Windows Update: {resultado_detalhes.stdout}")
        else:
            logging.error("Erro ao obter as propriedades do serviço Windows Update.")

    @staticmethod
    def ativar_servico_windows_update():
        """Tenta ativar e iniciar o serviço Windows Update."""
        logging.info("Verificando o status do serviço Windows Update...")

        # Verifica o status do serviço 'wuauserv'
        comando_verificar_status = COMANDO_POWERSHELL + [f"Get-Service -Name wuauserv | Select-Object -ExpandProperty Status"]
        resultado_status = executar_comando(comando_verificar_status, "Erro ao verificar o status do serviço Windows Update")

        if resultado_status and "Stopped" in resultado_status.stdout:
            logging.info("O serviço Windows Update está parado. Iniciando o serviço...")

            # Alterar o tipo de inicialização para automático
            executar_e_logar(
                COMANDO_POWERSHELL + [f"Set-Service -Name wuauserv -StartupType Automatic"],
                "Erro ao alterar o tipo de inicialização para Automático"
            ) # Falha ao alterar o tipo de inicialização

            # Tentar iniciar o serviço novamente
            comando_iniciar_servico = COMANDO_POWERSHELL + [f"Start-Service -Name wuauserv"]
            resultado_iniciar = executar_comando(comando_iniciar_servico, "Erro ao iniciar o serviço Windows Update")

            if resultado_iniciar and resultado_iniciar.returncode == 0:
                logging.info("Serviço Windows Update iniciado com sucesso.")
                return True  # Serviço iniciado corretamente
            else:
                logging.error(f"Falha ao iniciar o serviço Windows Update. Código de retorno: {resultado_iniciar.returncode}. Saída: {resultado_iniciar.stderr}")
                return False  # Falha ao iniciar o serviço

        elif "Running" in resultado_status.stdout:
            logging.info("O serviço Windows Update já está em execução.")
            return True

        else:
            logging.error("Falha ao verificar o status do serviço Windows Update.")
            return False

    @staticmethod
    def verificar_dependencias_windows_update():
        """Verifica dependências do serviço Windows Update."""

        comando_verificar_dependencias = COMANDO_POWERSHELL + [
            "Get-Service -Name wuauserv | Select-Object -ExpandProperty DependentServices | Format-List"
        ]
        resultado_dependencias = executar_comando(comando_verificar_dependencias, "Erro ao verificar dependências do serviço Windows Update")

        if resultado_dependencias and resultado_dependencias.returncode == 0:
           logging.info(f"Dependências verificadas com sucesso. {resultado_dependencias.stdout.strip() or 'Nenhuma dependência encontrada.'}")
        else:
           logging.error(f"Erro ao verificar as dependências do serviço Windows Update. "
                         f"Código de retorno: {resultado_dependencias.returncode}. "
                         f"Saída: {resultado_dependencias.stdout or 'Nenhuma saída'}. "
                         f"Erro: {resultado_dependencias.stderr or 'Nenhum erro capturado'}")

    @staticmethod
    def verificar_e_instalar_atualizacoes():
        """Verifica e instala atualizações do Windows, se necessário."""
        logging.info("Iniciando o processo de verificação de atualizações do Windows...")

        # Verificar dependências do serviço Windows Update
        ConfiguradorWindows.verificar_dependencias_windows_update()

        # Primeiro, garantir que o serviço Windows Update está ativo
        if not ConfiguradorWindows.ativar_servico_windows_update():
            logging.error("Não foi possível iniciar ou reiniciar o serviço Windows Update. Abortando a verificação de atualizações.")
            return  # Se o serviço não iniciar, o processo deve ser abortado.

        logging.info("Verificando atualizações pendentes...")

        # Verificar se o provedor NuGet está instalado
        comando_instalar_nuget = COMANDO_POWERSHELL + [
            "Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -Confirm:$false -Verbose | Out-String"
        ]
        resultado_instalacao_nuget = executar_comando(comando_instalar_nuget, "Erro ao instalar o provedor NuGet")

        if not resultado_instalacao_nuget and resultado_instalacao_nuget.returncode == 0:
            logging.error("Erro ao instalar o provedor NuGet.")
            return

        # Verificar se o módulo PSWindowsUpdate já está instalado
        comando_verificar_modulo = COMANDO_POWERSHELL + ["Get-Module -ListAvailable -Name PSWindowsUpdate"]
        resultado = executar_comando(comando_verificar_modulo, "Erro ao verificar o módulo PSWindowsUpdate")

        if not resultado and "PSWindowsUpdate" in resultado.stdout:
            logging.info("Módulo PSWindowsUpdate não encontrado. Tentando instalar...")

            # Tentar instalar o módulo PSWindowsUpdate
            comando_instalar_modulo = COMANDO_POWERSHELL + [
                "Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -Confirm:$false -SkipPublisherCheck"
            ]
            resultado_instalacao_modulo = executar_comando(comando_instalar_modulo, "Erro ao instalar o módulo PSWindowsUpdate")

            if not resultado_instalacao_modulo and resultado_instalacao_modulo.returncode == 0:
                logging.error(f"Erro ao instalar o módulo PSWindowsUpdate: {resultado_instalacao_modulo.stderr}")
                return

        # Importar o módulo PSWindowsUpdate explicitamente após a instalação
        comando_importar_modulo = COMANDO_POWERSHELL + ["Import-Module PSWindowsUpdate -Force"]
        resultado_importacao = executar_comando(comando_importar_modulo, "Erro ao importar o módulo PSWindowsUpdate")

        if resultado_importacao and resultado_importacao.returncode == 0:
            logging.info
        else:
            logging.error("Erro ao importar o módulo PSWindowsUpdate.")
            return

        # Verificar se há atualizações pendentes
        comando_verificar_atualizacoes = COMANDO_POWERSHELL + [
            "Get-WindowsUpdate -AcceptAll -ForceDownload -ForceInstall -ErrorAction SilentlyContinue"
        ]
        resultado_verificacao = executar_comando(comando_verificar_atualizacoes, "Erro ao verificar atualizações do Windows")

        if resultado_verificacao and resultado_verificacao.returncode == 0:
            if resultado_verificacao.stdout.strip():
                logging.info(f"Atualizações verificadas com sucesso: {resultado_verificacao.stdout}")
                ConfiguradorWindows.instalar_atualizacoes_pendentes()
            else:
                logging.info("Nenhuma atualização pendente foi encontrada.")
        else:
            logging.error(f"Erro ao verificar atualizações ou saída inesperada: {resultado_verificacao.stderr if resultado_verificacao else 'Nenhum resultado.'}")

    @staticmethod
    def instalar_atualizacoes_pendentes():
        """Instala todas as atualizações pendentes do Windows."""
        logging.info("Instalando atualizações pendentes do Windows...")

        comando_instalar_atualizacoes = COMANDO_POWERSHELL + ["Install-WindowsUpdate -AcceptAll -ErrorAction SilentlyContinue"]
        resultado_instalacao = executar_comando(comando_instalar_atualizacoes, "Erro ao instalar atualizações do Windows")

        if not resultado_instalacao and resultado_instalacao.returncode == 0:
            logging.error(f"Erro ao instalar atualizações do Windows: {resultado_instalacao.stderr}")

    @staticmethod
    def desativar_servicos_do_windows():
        """Desativa serviços como SysMain e Windows Update."""
        servicos = {
            "SysMain": "Serviço SysMain",
            "wuauserv": "Serviço Windows Update"
        }
        for nome_servico, descricao_servico in servicos.items():
            logging.info(f"Desativando {descricao_servico}...")
            comando_parar = COMANDO_POWERSHELL + [f"Stop-Service -Name {nome_servico} -Force; Set-Service -Name {nome_servico} -StartupType Disabled"]
            executar_e_logar(comando_parar, f"Erro ao desativar o {descricao_servico}")
            logging.info(f"{descricao_servico} foi desativado com sucesso.")

    @staticmethod
    def abrir_opcoes_de_desempenho():
        """Abre a janela 'Opções de Desempenho'."""
        try:
            logging.info("Abrindo janela de Opções de Desempenho...")
            app = Application().start('SystemPropertiesPerformance.exe')
            time.sleep(3)  # Aguarda a janela carregar
            return app
        except Exception as e:
            logging.error(f"Erro ao abrir a janela de Opções de Desempenho: {e}")
            return None

    @staticmethod
    def configurar_melhor_desempenho(app):
        """Seleciona 'Ajustar para obter um melhor desempenho' e ativa opções específicas."""
        try:
            # Conectar à janela principal
            dlg = app.window(title="Opções de Desempenho")
            dlg.wait("visible", timeout=10)
            #logging.info("Janela 'Opções de Desempenho' aberta e visível.")

            # Selecionar a opção "Ajustar para obter um &melhor desempenho"
            #logging.info("Selecionando a opção 'Ajustar para obter um melhor desempenho'...")
            radio_button = dlg.child_window(title="Ajustar para obter um &melhor desempenho", class_name="Button")
            if not radio_button.is_checked():
                radio_button.click()

            # Localizar o controle Tree1
            #logging.info("Explorando opções dentro de 'Tree1'...")
            tree = dlg.child_window(class_name="SysTreeView32")
            if tree.exists():
                for i in range(tree.item_count()):
                    item = tree.get_item([i])
                    #logging.info(f"Opção encontrada: {item.text()}")

                    # Verificar se é uma das opções a serem ativadas
                    if item.text() in ["Usar fontes de tela com cantos arredondados",
                                       "Usar sombras subjacentes para rótulos de ícones na área de trabalho"]:
                        #logging.info(f"Processando opção: {item.text()}")

                        # Simular clique apenas se a opção não estiver ativada
                        if not item.is_checked():
                            #logging.info(f"Ativando opção: {item.text()}")
                            item.click_input()  # Clique no item
                            #logging.info(f"Opção ativada: {item.text()}")
                        else:
                            logging.info(f"A opção já está ativada: {item.text()}")
            else:
                logging.error("O elemento Tree1 não foi encontrado.")

            # Clicar em "Ap&licar"
            #logging.info("Clicando no botão 'Ap&licar'...")
            aplicar_button = dlg.child_window(title="Ap&licar", class_name="Button")
            if aplicar_button.exists():
                aplicar_button.click()

            # Clicar em "OK"
            #logging.info("Clicando no botão 'OK'...")
            ok_button = dlg.child_window(title="OK", class_name="Button")
            if ok_button.exists():
                ok_button.click()

            logging.info("Configuração de 'Melhor Desempenho' aplicada com ajustes específicos.")
            return True
        except AttributeError as ae:
            logging.error(f"Erro ao acessar os itens da árvore: {ae}")
            return False
        except Exception as e:
            logging.error(f"Erro ao configurar 'Melhor Desempenho': {e}")
            return False

    @staticmethod
    def configurar_aparencia_desempenho():
        """Método principal para configurar 'Melhor Desempenho'."""
        app = ConfiguradorWindows.abrir_opcoes_de_desempenho()
        if app:
            sucesso = ConfiguradorWindows.configurar_melhor_desempenho(app)
            if sucesso:
                logging.info("Configuração concluída com sucesso.")
            else:
                logging.error("Falha ao aplicar a configuração de desempenho.")
        else:
            logging.error("Falha ao abrir a janela de configurações.")

    @staticmethod
    def verificar_versao_do_windows():
        """Verifica a versão do Windows e registra no log."""
        try:
            versao = platform.win32_ver()
            edicao = platform.win32_edition()
            logging.info(f"Versão do Windows: {versao[0]} {versao[1]} {versao[2]}, Edição: {edicao}")
        except Exception as e:
            logging.error(f"Erro ao verificar a versão do Windows: {e}")

# Função para emitir som de conclusão
def emitir_som_conclusao():
    """Emite um som de conclusão."""
    try:
        # Determina o caminho correto para o arquivo dentro do executável
        if getattr(sys, 'frozen', False):  # Caso seja executável PyInstaller
            base_path = sys._MEIPASS
        else:  # Durante execução do script no ambiente local
            base_path = os.path.dirname(__file__)

        caminho_audio = os.path.join(base_path, 'audio_conclusao.wav')

        if not os.path.exists(caminho_audio):
            print(f"Arquivo de som não encontrado: {caminho_audio}")
            return

        # Reproduz o som usando winsound
        winsound.PlaySound(caminho_audio, winsound.SND_FILENAME | winsound.SND_ASYNC)

    except Exception as e:
        print(f"Erro ao reproduzir som de conclusão: {e}")

# Função para exibir mensagem final e esperar interação do usuário
def esperar_usuario():
    try:
        print("\n### O script foi concluído com sucesso.")
        print("### Verifique se tudo foi realizado corretamente.")
        print("### Pressione Enter para encerrar...")
        input()
        logging.info("Usuário pressionou Enter, encerrando o script.")
    except EOFError:
        logging.info("O script foi concluído, mas não conseguiu capturar a interação de input do usuário.")
    except Exception as e:
        logging.error(f"Erro ao capturar input do usuário: {e}")
    finally:
        print("Encerrando o script em 3 segundos...")
        time.sleep(3)

# Função principal
def executar_script_completo():
    """
    Executa todas as etapas do script original de forma sequencial.
    """
    try:
        # Verificar versão do Windows
        ConfiguradorWindows.verificar_versao_do_windows()

        # Verificação e ativação do Windows
        AtivadorWindows.verificar_e_ativar_windows()

        # Verificar e instalar Chocolatey
        if not InstaladorSoftware.verificar_instalacao_do_chocolatey():
            logging.error("Falha ao instalar o Chocolatey.")
            return

        # Ativa o Office
        ativar_office()
        logging.info("Ativando pacote Office...")   

        # Remover completamente o Chocolatey após todas as instalações
        InstaladorSoftware.remover_chocolatey()

        # Verificar e instalar atualizações do Windows
        ConfiguradorWindows.verificar_e_instalar_atualizacoes()
        logging.info("Atualizações verificadas e instaladas com sucesso.")

        # Desativar serviços do Windows
        ConfiguradorWindows.desativar_servicos_do_windows()
        logging.info("Serviços desativados com sucesso.")

        # Configurar aparência e desempenho usando pyautogui
        ConfiguradorWindows.configurar_aparencia_desempenho()
        logging.info("Configuração de aparência e desempenho concluída.")

        # Exibir mensagem final para o usuário e esperar entrada
        esperar_usuario()
        logging.info("Finalização do script realizada com sucesso.")

    except Exception as e:
        logging.error(f"Erro ao executar o script completo: {e}")

# Interação do Usuário
def menu_interativo():
    """Apresenta um menu inicial para o usuário escolher a funcionalidade desejada."""
    while True:
        print("\nO que você deseja fazer?")
        print("1. Executar o script completo")
        print("2. Ativar o Windows")
        print("3. Ativar o Office")
        print("4. Instalar softwares")
        print("5. Configurações de desempenho")
        print("6. Sair")
        opcao = input("Escolha uma opção (1-6): ")

        if opcao == "1":
            executar_script_completo()  # Chama a função completa
            emitir_som_conclusao()

        elif opcao == "2":
            AtivadorWindows.verificar_e_ativar_windows()  # Chama a função que ativa o Windows
            emitir_som_conclusao()

        elif opcao == "3":
            ativar_office()  # Chama a função que ativa o pacote office
            emitir_som_conclusao()

        elif opcao == "4":
            menu_softwares()  # Chama o submenu de softwares

        elif opcao == "5":
            menu_desempenho()  # Chama o submenu de desempenho

        elif opcao == "6":
            print("Encerrando...")
            time.sleep(3)
            break

        else:
            print("Opção inválida. Por favor, escolha novamente.")

# Submenu Softwares
def menu_softwares():
    """
    Submenu para instalação de softwares.
    """
    # Lista de softwares disponíveis na ordem desejada
    softwares_disponiveis = [
        software for software in CONFIGURACAO_DE_SOFTWARE.keys()
        if software not in ['Word', 'Excel', 'PowerPoint']
    ]
    
    while True:
        print("\nInstalação de Softwares:")
        print("1. Instalar todos os softwares")
        print("2. Escolher um software específico")
        print("3. Voltar")
        sub_opcao = input("Escolha uma opção (1-3): ")

        if sub_opcao == "1":
            # Instalar todos os softwares disponíveis
            for software in softwares_disponiveis:
                InstaladorSoftware.instalar_software(software)
            emitir_som_conclusao()

        elif sub_opcao == "2":
            while True:
                print("\nSoftwares disponíveis:")
                for idx, software in enumerate(softwares_disponiveis, start=1):
                    print(f"{idx}. {software}")
                print(f"{len(softwares_disponiveis) + 1}. Voltar")
                software_idx = int(input(f"Escolha o número do software ou volte (1-{len(softwares_disponiveis) + 1}): ")) - 1

                if 0 <= software_idx < len(softwares_disponiveis):
                    software_escolhido = softwares_disponiveis[software_idx]
                    InstaladorSoftware.instalar_software(software_escolhido)
                    emitir_som_conclusao()

                elif software_idx == len(softwares_disponiveis):  # Opção "Voltar"
                    break
                else:
                    print("Opção inválida. Por favor, tente novamente.")

        elif sub_opcao == "3":
            break  # Voltar ao menu principal

        else:
            print("Opção inválida. Por favor, tente novamente.")

# Submenu desempenho
def menu_desempenho():
    """
    Submenu para configurações de desempenho.
    """
    while True:
        print("\nConfigurações de Desempenho:")
        print("1. Desativar serviços do Windows")
        print("2. Configurar aparência para melhor desempenho")
        print("3. Voltar")
        desempenho_opcao = input("Escolha uma opção (1-3): ")

        if desempenho_opcao == "1":
            ConfiguradorWindows.desativar_servicos_do_windows()
            emitir_som_conclusao()

        elif desempenho_opcao == "2":
            ConfiguradorWindows.configurar_aparencia_desempenho()
            emitir_som_conclusao()

        elif desempenho_opcao == "3":
            break  # Voltar ao menu principal

        else:
            print("Opção inválida. Por favor, tente novamente.")

def main():
    """
    Ponto de entrada principal para o script.
    """
    try:
        elevate()  # Eleva permissões para executar como administrador
        menu_interativo()

    except Exception as e:
        logging.error(f"Erro inesperado: {e}")

if __name__ == "__main__":
    main()
