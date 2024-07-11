import subprocess
import sys
import cv2
import numpy as np
import pyautogui as py
import pyperclip
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

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
        "PIL": "Pillow",
        "pyautogui": "pyautogui",
        "cv2": "opencv-python",
        "pyperclip": "pyperclip",
        "selenium": "selenium",
        "webdriver_manager": "webdriver-manager",
        "numpy": "numpy",
        "pandas": "pandas"
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

def find_image_on_screen(image_path, confidence=0.7):
    """
    Localiza uma imagem na tela com uma certa confiança.

    Args:
        image_path (str): Caminho para a imagem a ser encontrada.
        confidence (float): Nível de confiança para a correspondência.

    Returns:
        tuple: Coordenadas (x, y) do centro da imagem encontrada, ou None se não encontrada.
    """
    try:
        screenshot = py.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if template is None:
            raise Exception(f"Falha ao ler a imagem do arquivo: {image_path}")
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        result = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        if max_val >= confidence:
            template_height, template_width = template_gray.shape
            center_x = max_loc[0] + template_width // 2
            center_y = max_loc[1] + template_height // 2
            return center_x, center_y
        else:
            print(f"Imagem não encontrada com confiança suficiente. Confiança: {max_val}")
            return None
    except Exception as e:
        print(f"Erro ao encontrar a imagem na tela: {e}")
        return None

def click_on_image(image_path, confidence=0.7):
    """
    Clica na imagem localizada na tela com uma certa confiança.

    Args:
        image_path (str): Caminho para a imagem a ser clicada.
        confidence (float): Nível de confiança para a correspondência.

    Raises:
        Exception: Se a imagem não for encontrada.
    """
    location = find_image_on_screen(image_path, confidence)
    if location:
        py.moveTo(location)
        time.sleep(2)
        py.click(clicks=2)
        time.sleep(2)
        print(f"Clicou na posição {location}")
    else:
        raise Exception(f"Imagem {image_path} não encontrada com confiança {confidence}")

class WebAutomation:
    """
    Classe para automação de tarefas web usando Selenium e pyautogui.
    """
    def __init__(self):
        chrome_driver_path = ChromeDriverManager().install()
        self.nav = webdriver.Chrome(service=Service(chrome_driver_path))
    
    def wait_for_element(self, xpath, timeout=20):
        """
        Espera um elemento ser visível na página.

        Args:
            xpath (str): XPath do elemento.
            timeout (int): Tempo máximo de espera em segundos.

        Returns:
            WebElement: Elemento encontrado.
        """
        return WebDriverWait(self.nav, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def send_keys(self, xpath, keys):
        """
        Envia teclas para um elemento.

        Args:
            xpath (str): XPath do elemento.
            keys (str): Teclas a serem enviadas.
        """
        element = self.wait_for_element(xpath)
        element.clear()
        element.send_keys(keys)
    
    def click(self, xpath):
        """
        Clica em um elemento.

        Args:
            xpath (str): XPath do elemento.
        """
        element = self.wait_for_element(xpath)
        element.click()
    
    def verify_if_element_exists(self, xpath):
        """
        Verifica se um elemento existe na página.

        Args:
            xpath (str): XPath do elemento.

        Returns:
            bool: True se o elemento existe, False caso contrário.
        """
        try:
            self.nav.find_element(By.XPATH, xpath)
            return True
        except Exception:
            return False

    def open(self):
        """
        Abre o domínio web.
        """
        print("Abrindo o domínio web...")
        self.nav.get("https://www.dominioweb.com.br/")
        self.nav.maximize_window()
        self.wait_for_element('//*[@id="trid-auth"]/form/label[1]/span[2]/input')

    def fechar_mensagem_chrome(self):
        """
        Fecha a mensagem do Chrome.
        """
        py.press('esc')
        time.sleep(1)

    def login(self, user, password):
        """
        Faz login no domínio.

        Args:
            user (str): Nome de usuário.
            password (str): Senha.
        """
        try:
            print("Iniciando processo de login...")
            self.send_keys('//*[@id="trid-auth"]/form/label[1]/span[2]/input', user)
            self.send_keys('//*[@id="trid-auth"]/form/label[2]/span[2]/input', password)
            self.click('//*[@id="enterButton"]')
            time.sleep(5)  # Aumentado o tempo de espera
            
            # Pressionar TAB, aguardar 1 segundo, pressionar TAB novamente, aguardar 1 segundo, e pressionar ENTER usando PyAutoGUI
            py.press('tab')
            time.sleep(1)
            py.press('tab')
            time.sleep(1)
            py.press('enter')

            print("Login e navegação bem-sucedidos")
            
        except Exception as e:
            print(f"Erro durante o login: {e}")
            raise


    def close(self):
        """
        Fecha o navegador.
        """
        time.sleep(30)  # Espera de 1 minuto antes de fechar o navegador
        self.nav.quit()

    def operar_aplicacao(self, texto_importacao, servico_image, app_user, app_password):
        """
        Opera a aplicação AppController.exe.

        Args:
            texto_importacao (str): Texto a ser digitado.
            servico_image (str): Caminho para a imagem do serviço.
            app_user (str): Nome de usuário do aplicativo.
            app_password (str): Senha do aplicativo.
        """
        button_image = os.path.join(os.getcwd(), 'img', 'escrita_fiscal_button.png')
        max_wait_time = 60  # Aumentado o tempo de espera
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            windows = py.getWindowsWithTitle("Lista de Programas")
            if windows:
                break
            print("Esperando pela janela 'Lista de Programas'...")
            time.sleep(1)
        else:
            print("A janela 'Lista de Programas' não foi encontrada. Tentando verificar outras janelas visíveis.")
            windows = py.getAllWindows()
            for window in windows:
                print(f"Janela encontrada: {window.title}")
            raise Exception("A janela 'Lista de Programas' não foi encontrada")
        time.sleep(5)  # Aumentado o tempo de espera
        try:
            click_on_image(button_image)
        except Exception as e:
            print(f"Erro ao clicar na imagem 'Escrita Fiscal': {e}")
        time.sleep(10)  # Espera antes de inserir credenciais
        self.__inserir_credenciais(app_user, app_password)
        self.__verificar_arquivos(texto_importacao, servico_image)

    def __inserir_credenciais(self, app_user, app_password):
        """
        Insere as credenciais na aplicação.

        Args:
            app_user (str): Nome de usuário do aplicativo.
            app_password (str): Senha do aplicativo.
        """
        time.sleep(2)
        py.write(app_user)  # Garantir que está escrevendo uma string
        py.press('tab')
        py.write(app_password)  # Garantir que está escrevendo uma string
        py.press('tab', presses=2)
        py.press('enter')

    def __verificar_arquivos(self, texto_importacao, servico_image):
        """
        Verifica se a imagem dos arquivos está presente.

        Args:
            texto_importacao (str): Texto a ser digitado.
            servico_image (str): Caminho para a imagem do serviço.
        """
        max_wait_time = 60
        start_time = time.time()
        arquivos_image = os.path.join(os.getcwd(), 'img', 'arquivos.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(arquivos_image, confidence=0.7)
            if location:
                py.click(location)
                self.__verificar_rotinas_automaticas(texto_importacao, servico_image)
                return
            time.sleep(1)
        raise Exception("A imagem 'arquivos.png' não foi encontrada")

    def __verificar_rotinas_automaticas(self, texto_importacao, servico_image):
        """
        Verifica se a imagem das rotinas automáticas está presente.

        Args:
            texto_importacao (str): Texto a ser digitado.
            servico_image (str): Caminho para a imagem do serviço.
        """
        rotinas_automaticas_image = os.path.join(os.getcwd(), 'img', 'RotinasAutomaticas.png')
        max_wait_time = 60
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(rotinas_automaticas_image, confidence=0.7)
            if location:
                py.click(location)
                self.__clicar_botao_novo_arquivos(texto_importacao, servico_image)
                return
            time.sleep(0.5)
        raise Exception("A imagem 'RotinasAutomaticas.png' não foi encontrada")

    def __clicar_botao_novo_arquivos(self, texto_importacao, servico_image):
        """
        Aguarda o botão 'botaoNovoArquivos.png' estar visível e clica nele.

        Args:
            texto_importacao (str): Texto a ser digitado.
            servico_image (str): Caminho para a imagem do serviço.
        """
        max_wait_time = 60
        start_time = time.time()
        botao_novo_arquivos_image = os.path.join(os.getcwd(), 'img', 'botaoNovoArquivos.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(botao_novo_arquivos_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou no botão 'Novo Arquivos' localizado em {location}")
                time.sleep(1)  # Aguarda meio segundo para a nova janela carregar
                py.press('tab')  # Pressiona a tecla Tab
                time.sleep(0.5)  # Aguarda meio segundo
                self.__digitar_com_pyperclip(texto_importacao)  # Digita o texto de importação
                self.__clicar_apuracao_importacao(servico_image)
                return
            time.sleep(1)
        raise Exception("A imagem 'botaoNovoArquivos.png' não foi encontrada")

    def __clicar_apuracao_importacao(self, servico_image):
        """
        Clica nas imagens 'Apuracao.png', 'Importacao.png', 'ImportarRegistros.png', 'ImportarNotas.png', 'NFeABRASF.png' e 'tresPontosABRASF.png'.

        Args:
            servico_image (str): Caminho para a imagem do serviço.
        """
        max_wait_time = 60
        start_time = time.time()
        
        # Encontrar e clicar na imagem Apuracao.png
        apuracao_image = os.path.join(os.getcwd(), 'img', 'Apuracao.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(apuracao_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'Apuracao.png' localizada em {location}")
                break
            time.sleep(1)
        else:
            raise Exception("A imagem 'Apuracao.png' não foi encontrada")
        
        time.sleep(1)  # Espera um segundo antes de clicar na próxima imagem

        # Encontrar e clicar na imagem Importacao.png
        importacao_image = os.path.join(os.getcwd(), 'img', 'Importacao.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(importacao_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'Importacao.png' localizada em {location}")
                break
            time.sleep(1)
        else:
            raise Exception("A imagem 'Importacao.png' não foi encontrada")
        
        time.sleep(1)  # Espera meio segundo antes de clicar na próxima imagem

        # Encontrar e clicar na imagem ImportarRegistros.png
        importar_registros_image = os.path.join(os.getcwd(), 'img', 'ImportarRegistros.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(importar_registros_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'ImportarRegistros.png' localizada em {location}")
                break
            time.sleep(1)
        else:
            raise Exception("A imagem 'ImportarRegistros.png' não foi encontrada")
        
        time.sleep(1)  # Espera meio segundo antes de clicar na próxima imagem

        # Encontrar e clicar na imagem ImportarNotas.png
        importar_notas_image = os.path.join(os.getcwd(), 'img', 'ImportarNotas.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(importar_notas_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'ImportarNotas.png' localizada em {location}")
                break
            time.sleep(1)
        else:
            raise Exception("A imagem 'ImportarNotas.png' não foi encontrada")
        
        time.sleep(1)  # Espera meio segundo antes de clicar na próxima imagem

        # Encontrar e clicar na imagem NFeABRASF.png
        nfe_abrasf_image = os.path.join(os.getcwd(), 'img', 'NFeABRASF.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(nfe_abrasf_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'NFeABRASF.png' localizada em {location}")
                break
            time.sleep(1)
        else:
            raise Exception("A imagem 'NFeABRASF.png' não foi encontrada")

        time.sleep(0.5)  # Espera meio segundo antes de clicar na próxima imagem

        # Encontrar e clicar na imagem tresPontosABRASF.png
        tres_pontos_abrasf_image = os.path.join(os.getcwd(), 'img', 'tresPontosABRASF.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(tres_pontos_abrasf_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'tresPontosABRASF.png' localizada em {location}")
                time.sleep(1)
                self.__clicar_e_inserir_caminho(servico_image)
                break
            time.sleep(2)
        else:
            raise Exception("A imagem 'tresPontosABRASF.png' não foi encontrada")

    def __clicar_e_inserir_caminho(self, servico_image):
        """
        Clica na área especificada pelas coordenadas obtidas pelo Inspect, insere o caminho especificado,
        clica na imagem 'SalvarRelatorio.png', clica em outra área, insere o mesmo caminho novamente,
        e clica nas imagens 'SemprePastaCompetencia.png', 'ServicoTomado.png' ou 'ServicoPrestado.png' e 'OK.png'.

        Args:
            servico_image (str): Caminho para a imagem do serviço.
        """
        caminho_texto = 'M:\\Projeto GoMind\\Projeto_GoMind\\ImportacaoEmLotes'
        max_wait_time = 60
        start_time = time.time()

        # Coordenadas obtidas pelo Inspect
        x, y = 482, 258

        print(f"Clicando nas coordenadas ({x}, {y})")

        # Clique nas coordenadas especificadas
        while time.time() - start_time < max_wait_time:
            try:
                py.moveTo(x, y)
                time.sleep(1)  # Aguarda 1 segundo antes de clicar
                py.click()
                time.sleep(1)  # Aguarda 1 segundo após clicar
                pyperclip.copy(caminho_texto)
                py.hotkey('ctrl', 'v')
                time.sleep(2)  # Aguarda 2 segundos para garantir que o texto seja colado
                print("Digitou o caminho nas coordenadas especificadas")

                # Procura e clica na imagem 'SalvarRelatorio.png'
                salvar_relatorio_image = os.path.join(os.getcwd(), 'img', 'SalvarRelatorio.png')
                salvar_location = find_image_on_screen(salvar_relatorio_image, confidence=0.7)
                if salvar_location:
                    print(f"Imagem 'SalvarRelatorio.png' encontrada em: {salvar_location}")
                    py.click(salvar_location)
                    print(f"Clicou na imagem 'SalvarRelatorio.png' localizada em {salvar_location}")

                    # Nova etapa: clicar na área da tela especificada
                    novo_x, novo_y = 497, 284
                    time.sleep(1)  # Aguarda 1 segundo antes de clicar na nova área
                    py.moveTo(novo_x, novo_y)
                    time.sleep(1)  # Aguarda 1 segundo antes de clicar
                    py.click()
                    time.sleep(1)  # Aguarda 1 segundo após clicar
                    py.hotkey('ctrl', 'v')
                    time.sleep(2)  # Aguarda 2 segundos para garantir que o texto seja colado
                    print(f"Digitou o caminho nas novas coordenadas ({novo_x}, {novo_y})")

                    # Clica na imagem 'SemprePastaCompetencia.png'
                    sempre_pasta_competencia_image = os.path.join(os.getcwd(), 'img', 'SemprePastaCompetencia.png')
                    sempre_pasta_competencia_location = find_image_on_screen(sempre_pasta_competencia_image, confidence=0.7)
                    if sempre_pasta_competencia_location:
                        time.sleep(1)  # Aguarda 1 segundo antes de clicar
                        py.click(sempre_pasta_competencia_location)
                        print(f"Clicou na imagem 'SemprePastaCompetencia.png' localizada em {sempre_pasta_competencia_location}")

                        # Clica na imagem do serviço (ServicoTomado.png ou ServicoPrestado.png)
                        servico_location = find_image_on_screen(servico_image, confidence=0.7)
                        if servico_location:
                            time.sleep(1)  # Aguarda 1 segundo antes de clicar
                            py.click(servico_location)
                            print(f"Clicou na imagem do serviço localizada em {servico_location}")

                            # Nova etapa: clica na imagem 'OK.png'
                            ok_image = os.path.join(os.getcwd(), 'img', 'OK.png')
                            ok_location = find_image_on_screen(ok_image, confidence=0.7)
                            if ok_location:
                                time.sleep(1)  # Aguarda 1 segundo antes de clicar
                                py.click(ok_location)
                                print(f"Clicou na imagem 'OK.png' localizada em {ok_location}")
                                return
                            else:
                                print("A imagem 'OK.png' não foi encontrada")
                                raise Exception("A imagem 'OK.png' não foi encontrada")
                        else:
                            print("A imagem do serviço não foi encontrada")
                            raise Exception("A imagem do serviço não foi encontrada")
                    else:
                        print("A imagem 'SemprePastaCompetencia.png' não foi encontrada")
                        raise Exception("A imagem 'SemprePastaCompetencia.png' não foi encontrada")
                else:
                    print("A imagem 'SalvarRelatorio.png' não foi encontrada")
                    raise Exception("A imagem 'SalvarRelatorio.png' não foi encontrada")
            except Exception as e:
                print(f"Erro ao clicar nas coordenadas ou colar o caminho: {e}")

            time.sleep(1)

        print("Não foi possível clicar nas coordenadas ou encontrar a imagem 'SalvarRelatorio.png'")
        raise Exception("A imagem 'CaminhoTomado.png' não foi encontrada")

    def __digitar_com_pyperclip(self, text):
        """
        Utiliza pyperclip para copiar e colar texto, garantindo caracteres especiais.

        Args:
            text (str): Texto a ser copiado e colado.
        """
        pyperclip.copy(text)
        py.hotkey('ctrl', 'v')

def ler_credenciais(caminho_arquivo):
    """
    Lê as credenciais de um arquivo Excel.

    Args:
        caminho_arquivo (str): Caminho para o arquivo Excel.

    Returns:
        tuple: Credenciais (web_username, web_password, app_username, app_password).
    """
    df = pd.read_excel(caminho_arquivo, dtype=str)  # Forçar leitura como string
    web_username = df.loc[0, 'web_username']
    web_password = df.loc[0, 'web_password']
    app_username = df.loc[0, 'app_username']
    app_password = df.loc[0, 'app_password']
    return web_username, web_password, app_username, app_password

if verificar_dependencias():
    caminho_arquivo_credenciais = 'credentials.xlsx'
    web_username, web_password, app_username, app_password = ler_credenciais(caminho_arquivo_credenciais)
    
    web_automation = WebAutomation()
    web_automation.open()

    try:
        web_automation.login(web_username, web_password)
    except Exception as e:
        print(f"Erro ao tentar logar: {e}")

    try:
        web_automation.operar_aplicacao('Importação Entrada Tomado', os.path.join(os.getcwd(), 'img', 'ServicoTomado.png'), app_username, app_password)
    except Exception as e:
        print(f"Erro ao operar a aplicação: {e}")

    try:
        web_automation.close()
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")
