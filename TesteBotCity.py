import subprocess
import sys
import time
import os
from botcity.core import DesktopBot
from botcity.web import WebBot, Browser, By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def instalar_dependencia(pacote):
    """
    Instala o pacote especificado usando pip.

    Args:
        pacote (str): Nome do pacote a ser instalado.
    """
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
    except subprocess.CalledProcessError as e:
        print(f"Erro ao instalar o pacote {pacote}: {e}")
        return False
    return True

def verificar_dependencias():
    """
    Verifica se as dependências necessárias estão instaladas e as instala se necessário.
    """
    dependencias = {
        "botcity.core": "botcity-framework-core",
        "botcity.framework.core": "botcity-framework-core",
        "botcity.framework.web": "botcity-framework-web"
    }
    
    for modulo, pacote in dependencias.items():
        try:
            __import__(modulo)
            print(f"{modulo} está instalado corretamente.")
        except ImportError:
            print(f"{modulo} não está instalado. Instalando {pacote}...")
            if not instalar_dependencia(pacote):
                print(f"Falha ao instalar {pacote}.")
                return False
            try:
                __import__(modulo)
                print(f"{modulo} instalado com sucesso.")
            except ImportError:
                print(f"Falha ao importar {modulo} após instalação.")
                return False

    return True

class WebAutomation(WebBot):
    """
    Classe para automação de tarefas web usando BotCity WebBot.
    """
    def __init__(self):
        super().__init__()
        self.headless = False
        self.driver_path = ChromeDriverManager().install()
        self.driver = self.get_browser_driver(Browser.CHROME, self.driver_path)

    def wait_for_element(self, xpath, timeout=20):
        """
        Espera um elemento ser visível na página.

        Args:
            xpath (str): XPath do elemento.
            timeout (int): Tempo máximo de espera em segundos.

        Returns:
            WebElement: Elemento encontrado.
        """
        return WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def open(self, url):
        """
        Abre o domínio web.
        """
        self.browse(url, Browser.CHROME)
        self.driver.maximize_window()
        self.wait_for_element('//*[@id="trid-auth"]/form/label[1]/span[2]/input')

    def aceitar_plugin_dominio(self):
        """
        Aceita o plugin do domínio.
        """
        image_path = r'C:\Projeto GoMind\Projeto_GoMind\ImportacaoEmLotes\img\abrirPlugin_TRComputer.png'
        if not os.path.isfile(image_path):
            raise Exception(f"Arquivo de imagem não encontrado: {image_path}")
        time.sleep(10)
        self.fechar_mensagem_chrome()
        try:
            self.click_on_image(image_path)
        except Exception as e:
            print(f"Erro ao clicar na imagem do plugin: {e}")

    def fechar_mensagem_chrome(self):
        """
        Fecha a mensagem do Chrome.
        """
        self.send_keys(self.driver.find_element_by_tag_name('body'), 'esc')

    def login(self, user, password):
        """
        Faz login no domínio.

        Args:
            user (str): Nome de usuário.
            password (str): Senha.
        """
        try:
            self.send_keys('//*[@id="trid-auth"]/form/label[1]/span[2]/input', user)
            self.send_keys('//*[@id="trid-auth"]/form/label[2]/span[2]/input', password)
            self.click('//*[@id="enterButton"]')
            time.sleep(10)  # Aumentado o tempo de espera
            # Verifique se qualquer elemento indica um login bem-sucedido
            if self.verify_if_element_exists("//*[contains(text(), 'Dashboard')]") or \
               self.verify_if_element_exists("//*[contains(text(), 'Bem-vindo')]"):
                print("Login bem-sucedido")
            else:
                raise Exception("Login falhou")
        except Exception as e:
            print(f"Erro durante o login: {e}")
            raise

    def escolher_modulo(self, coluna=1, linha=1):
        """
        Escolhe o módulo no domínio.

        Args:
            coluna (int): Número da coluna do módulo.
            linha (int): Número da linha do módulo.
        """
        if coluna == 1 and linha == 1:
            self.key_right()
            self.key_left()
        else:
            for _ in range(coluna - 1):
                self.key_right()
            for _ in range(linha - 1):
                self.key_down()
        time.sleep(1)
        self.key_enter()
        self.send_hotkey("win", "d")

    def close(self):
        """
        Fecha o navegador.
        """
        self.driver.quit()

    def operar_aplicacao(self):
        """
        Opera a aplicação AppController.exe.
        """
        button_image = r'C:\Projeto GoMind\Projeto_GoMind\ImportacaoEmLotes\img\escrita_fiscal_button.png'
        max_wait_time = 60  # Aumentado o tempo de espera
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            windows = self.get_windows_with_title("Lista de Programas")
            if windows:
                break
            time.sleep(1)
        else:
            raise Exception("A janela 'Lista de Programas' não foi encontrada")
        time.sleep(5)  # Aumentado o tempo de espera
        try:
            self.click_on_image(button_image)
        except Exception as e:
            print(f"Erro ao clicar na imagem 'Escrita Fiscal': {e}")
        time.sleep(10)  # Espera antes de inserir credenciais
        self.__inserir_credenciais()
        self.__verificar_arquivos()

    def __inserir_credenciais(self):
        """
        Insere as credenciais na aplicação.
        """
        self.write('JOHN.DOMINI')
        self.press('tab')
        self.write('030201')
        self.press('tab', 2)
        self.press('enter')

    def __verificar_arquivos(self):
        """
        Verifica se a imagem dos arquivos está presente.
        """
        max_wait_time = 60
        start_time = time.time()
        arquivos_image = r'C:\Projeto GoMind\Projeto_GoMind\ImportacaoEmLotes\img\arquivos.png'
        while time.time() - start_time < max_wait_time:
            location = self.find_image_on_screen(arquivos_image, confidence=0.9)  # Aumentada a confiança mínima
            if location:
                self.click(location)
                self.__verificar_rotinas_automaticas()
                return
            time.sleep(1)
        raise Exception("A imagem 'arquivos.png' não foi encontrada")

    def __verificar_rotinas_automaticas(self):
        """
        Verifica se a imagem das rotinas automáticas está presente.
        """
        rotinas_automaticas_image = r'C:\Projeto GoMind\Projeto_GoMind\ImportacaoEmLotes\img\RotinasAutomaticas.png'
        max_wait_time = 60
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            location = self.find_image_on_screen(rotinas_automaticas_image, confidence=0.9)  # Aumentada a confiança mínima
            if location:
                self.click(location)
                return
            time.sleep(0.5)
        raise Exception("A imagem 'RotinasAutomaticas.png' não foi encontrada")

if verificar_dependencias():
    web_automation = WebAutomation()
    web_automation.open("https://www.dominioweb.com.br/")

    username = 'john.domini@grupodomini.com'
    password = 'J0hN@RPA!2024'

    try:
        web_automation.login(username, password)
    except Exception as e:
        print(f"Erro ao tentar logar: {e}")

    try:
        web_automation.escolher_modulo(coluna=2, linha=3)
    except Exception as e:
        print(f"Erro ao escolher o módulo: {e}")

    try:
        web_automation.operar_aplicacao()
    except Exception as e:
        print(f"Erro ao operar a aplicação: {e}")

    try:
        web_automation.close()
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")
