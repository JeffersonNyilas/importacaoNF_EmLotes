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
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

def instalar_dependencia(pacote):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
    except subprocess.CalledProcessError as e:
        print(f"Erro ao instalar o pacote {pacote}: {e}")
        return False
    return True

def verificar_dependencias():
    dependencias = {
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

def double_click_on_image(image_path, confidence=0.7):
    location = find_image_on_screen(image_path, confidence)
    if location:
        py.moveTo(location)
        time.sleep(1)
        py.doubleClick()  # Realiza o duplo clique
        time.sleep(1)
        print(f"Duplo clique na posição {location}")
    else:
        raise Exception(f"Imagem {image_path} não encontrada com confiança {confidence}")

class WebAutomation:
    def __init__(self):
        chrome_driver_path = ChromeDriverManager().install()
        self.nav = webdriver.Chrome(service=Service(chrome_driver_path))
    
    def wait_for_element(self, xpath, timeout=20):
        return WebDriverWait(self.nav, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def send_keys(self, xpath, keys):
        element = self.wait_for_element(xpath)
        element.clear()
        element.send_keys(keys)
    
    def click(self, xpath):
        element = self.wait_for_element(xpath)
        element.click()
    
    def verify_if_element_exists(self, xpath):
        try:
            self.nav.find_element(By.XPATH, xpath)
            return True
        except Exception:
            return False

    def open(self):
        print("Abrindo o domínio web...")
        self.nav.get("https://www.dominioweb.com.br/")
        self.nav.maximize_window()
        self.wait_for_element('//*[@id="trid-auth"]/form/label[1]/span[2]/input')

    def fechar_mensagem_chrome(self):
        py.press('esc')
        time.sleep(1)

    def login(self, user, password):
        try:
            print("Iniciando processo de login...")
            self.send_keys('//*[@id="trid-auth"]/form/label[1]/span[2]/input', user)
            self.send_keys('//*[@id="trid-auth"]/form/label[2]/span[2]/input', password)
            self.click('//*[@id="enterButton"]')
            time.sleep(5)
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
        time.sleep(30)
        self.nav.quit()

    def operar_aplicacao(self, app_username, app_password):
        button_image = os.path.join(os.getcwd(), 'img', 'escrita_fiscal_button.png')
        max_wait_time = 60
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
        time.sleep(5)
        try:
            double_click_on_image(button_image)
        except Exception as e:
            print(f"Erro ao clicar na imagem 'Escrita Fiscal': {e}")
        time.sleep(10)
        self.__inserir_credenciais(app_username, app_password)
        self.__verificar_arquivos()

    def __inserir_credenciais(self, app_user, app_password):
        time.sleep(2)
        py.write(app_user)  # Garantir que está escrevendo uma string
        py.press('tab')
        py.write(app_password)  # Garantir que está escrevendo uma string
        py.press('tab', presses=2)
        py.press('enter')

    def __verificar_arquivos(self):
        max_wait_time = 60
        start_time = time.time()
        arquivos_image = os.path.join(os.getcwd(), 'img', 'Movimentos.png')
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(arquivos_image, confidence=0.7)
            if location:
                py.click(location)
                self.__verificar_rotinas_automaticas()
                return
            time.sleep(1)
        raise Exception("A imagem 'Movimentos.png' não foi encontrada")

    def __verificar_rotinas_automaticas(self):
        rotinas_automaticas_image = os.path.join(os.getcwd(), 'img', 'RotinasAutomaticas2.png')
        max_wait_time = 60
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            location = find_image_on_screen(rotinas_automaticas_image, confidence=0.7)
            if location:
                py.click(location)
                print(f"Clicou na imagem 'RotinasAutomaticas.png' localizada em {location}")
                time.sleep(1)
                self.__clicar_novo_executar()
                return
            time.sleep(0.5)
        raise Exception("A imagem 'RotinasAutomaticas.png' não foi encontrada")

    def __clicar_novo_executar(self):
        time.sleep(1)
        novo_executar_image = os.path.join(os.getcwd(), 'img', 'NovoExecutar.png')
        double_click_on_image(novo_executar_image, confidence=0.7)

def ler_credenciais(caminho_arquivo):
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
        web_automation.operar_aplicacao(app_username, app_password)  # Passando as credenciais como argumento
    except Exception as e:
        print(f"Erro ao operar a aplicação: {e}")

    try:
        web_automation.close()
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")
