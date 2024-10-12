import os
import sys
import shutil
import zipfile
from tkinter import messagebox, Tk
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Função para exibir uma mensagem de erro em um alerta
def exibir_alerta(mensagem):
    root = Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    messagebox.showerror("Erro", mensagem)
    root.destroy()  # Fecha a janela tkinter após o alerta

# Função para exibir uma mensagem de sucesso em um alerta
def exibir_sucesso(mensagem):
    root = Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    messagebox.showinfo("Sucesso", mensagem)
    root.destroy()  # Fecha a janela tkinter após o alerta

# Função para exibir a mensagem para o usuário
def show_message():
    root = Tk()
    root.withdraw() # Esconde a janela principal do tkinter
    # Garantir que a janela de mensagem fique no topo
    root.attributes("-topmost", True)  # Define a janela como "topmost", garantindo foco
    messagebox.showinfo("Ação necessária", "Entre no site, insira o código de verificação enviado por email e clique em OK no navegador para continuar.")
    root.quit()  # Fecha a janela do Tkinter após o usuário clicar em OK

# Função extrair o chromedriver de um arquivo zip do local de origem
def extrair_chromedriver_local(caminho_zip, destino, temp_dir):
    try:
        # Cria o diretório temporário
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_zip):
            raise Exception(f"Arquivo {caminho_zip} não encontrado. Verifique o caminho e tente novamente.")

        # Faz uma cópia temporária do ZIP para preservar o original
        zip_temp = os.path.join(temp_dir, os.path.basename(caminho_zip))
        shutil.copy2(caminho_zip, zip_temp)

        # Extrai o arquivo da cópia temporária
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            print("Arquivo extraído para {}".format(temp_dir))

        # Procura pelo chromedriver.exe dentro de uma pasta temporária e move para o destino
        for root, dirs, files in os.walk(temp_dir):
            if "chromedriver.exe" in files:
                chromedriver_exe = os.path.join(root, "chromedriver.exe")
                shutil.move(chromedriver_exe, destino)
                print("Chromedriver copiado para: {}".format(destino))
                break
        else:
            # Caso o arquivo não seja encontrado, exibe um alerta e sai com código de erro 3
            exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
            sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip

        # Remove o arquivo zip após extração
        os.remove(zip_temp)
        print("Arquivo zip removido após extração.")

    except Exception as e:
        exibir_alerta("Falha ao extrair o chromedriver: {}".format(str(e)))
        sys.exit(2)  # Código de erro 2 para falha ao baixar e extrair arquivo

# Função para configurar e inicializar o chromedriver
def configurar_chromedriver(caminho_zip, chromedriver_path, temp_dir):
    # Verifica se o chromedriver já está disponível
    if not chromedriver_path.exists():
        print("Chromedriver não encontrado localmente. Extraindo do zip ...")
        extrair_chromedriver_local(caminho_zip, str(chromedriver_path), str(temp_dir))
    else:
        print("Chromedriver já está configurado.")

# Função para inicializar o webdriver
def inicializar_navegador(chromedriver_path, headless=False):
    try:
        # Configurar as opções do Chrome
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless") # Rodar o navegador sem interface gráfica (opcional)

        # Criar o objeto Service com o caminho do chromedriver
        service = Service(executable_path=chromedriver_path)
        # Inicializa o WebDriver
        navegador = webdriver.Chrome(service=service, options=chrome_options)
        print("Navegador inicializado com sucesso.")
        return navegador
    except Exception as e:
        print(f"Erro ao inicializar o navegador: {e}")
        exibir_alerta("Erro ao inicializar o navegador: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Definir os caminhos
    caminho_zip = r"C:\Users\erics\PycharmProjects\pythonProject\resources\chromedriver.zip"
    # Definição dos caminhos e variáveis
    chromedriver_path = Path("C:/Users/{}/AppData/Local/chromedriver/chromedriver.exe".format(os.getlogin()))
    temp_dir = Path(chromedriver_path).parent / 'temp'

    # Configurar o chromedriver
    configurar_chromedriver(caminho_zip, chromedriver_path, temp_dir)
    # Inicializar o navegador
    navegador = inicializar_navegador(chromedriver_path, headless=False)

    if navegador:
        # Executar ações no navegador (exemplo)
        navegador.get("https://www.example.com")
        exibir_sucesso("Navegador inicializado e site carregado.")
        navegador.quit()
