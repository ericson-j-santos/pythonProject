import os
import sys
import shutil
import urllib.request
import zipfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from tkinter import messagebox, Tk
from pathlib import Path

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

# Função para baixar e extrair o chromedriver de um arquivo zip do local de origem
# def baixar_e_extrair_chromedriver(url, destino, temp_dir):
def extrair_chromedriver_local(caminho_zip, destino, temp_dir):
    try:
        # Cria o diretório temporário
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_zip):
            raise Exception(f"Arquivo {caminho_zip} não encontrado. Verifique o caminho e tente novamente.")

        # # Caminho do arquivo zip temporário
        # zip_path = os.path.join(temp_dir, "chromedriver.zip")
        #
        # # Baixa o arquivo zip da pasta origem com timeout manual
        # with urllib.request.urlopen(url, timeout=60) as response:
        #     with open(zip_path, 'wb') as out_file:
        #         shutil.copyfileobj(response, out_file)
        # print("Arquivo baixado para: {}.".format(zip_path))
        #
        # # Verifica se o arquivo baixado não está vazio
        # if os.path.getsize(zip_path) == 0:
        #     raise Exception("Arquivo baixado está vazio ou corrompido.")

        # Extrai o arquivo zip
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            print("Arquivo extraído para {}".format(temp_dir))

        # Procura pelo chromedriver.exe dentro de uma pasta extraída
        for root, dirs, files in os.walk(temp_dir):
            if "chromedriver.exe" in files:
                chromedriver_exe = os.path.join(root, "chromedriver.exe")
                # Move o chromedriver para o destino final
                shutil.move(chromedriver_exe, destino)
                print("Chromedriver copiado para: {}".format(destino))
                break
        else:
            # Caso o arquivo não seja encontrado, exibe um alerta e sai com código de erro 3
            exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
            sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip

        # Remove o arquivo zip após extração
        os.remove(caminho_zip)
        print("Arquivo zip removido após extração.")

    # except urllib.error.URLError as e:
    #     exibir_alerta("Falha ao baxar o chromedriver: Verifique a URL ou sua conexão com a internet.")
    #     sys.exit(2)
    except Exception as e:
        exibir_alerta("Falha ao extrair o chromedriver: {}".format(str(e)))
        sys.exit(2)  # Código de erro 2 para falha ao baixar e extrair arquivo



# Função para inicializar o webdriver
def inicializar_navegador(chromedriver_path):
    try:

        # Debugging: Verificando o ambiente e o caminho do chromedriver
        # print("Python Executável:", sys.executable)
        # print("Versão do Python:", sys.version)
        # print("Caminho do chromedriver:", chromedriver_path)

        #Configurar as opções do Chrome
        chrome_options = Options()
        #chrome_options.add_argument("--headless") # Rodar o navegador sem interface gráfica (opcional)

        # Criar o objeto Service com o caminho do chromedriver
        service = Service(executable_path=chromedriver_path)

        # Inicializa o webdriver diretamente com o caminho do chromedriver
        navegador = webdriver.Chrome(service=service, options=chrome_options)  # Ajuste aqui, passando o caminho como argumento posicional

        # Define timeout explicitamente (opcional, se necessário)
        #navegador.set_page_load_timeout(10)  # Define timeout de 30 segundos para carregamento da página

        navegador.get("https://www.google.com.br")
        print("Navegador foi aberto com sucesso")
        navegador.quit()
    except Exception as e:
        exibir_alerta("Ocorreu um erro ao abrir o navegador: {}".format(str(e)))
        sys.exit(4)  # Código de erro 4 para falha ao inicializar navegador
    finally:
        if 'navegador' in locals():
            navegador.quit() # Garante que o navegador seja fechado

try:
    # Caminhos de configuração
    nome_usuario = os.getlogin() # Obtém o nome do usuário atual
    base_dir = Path("C:/Users/{}/AppData/Local/chromedriver".format(nome_usuario))  # Caminho para o diretório 'AppData\Local' do usuário
    chromedriver_path = base_dir / "chromedriver.exe"  # Caminho para a pasta onde o 'chromedriver.exe' será colocado
    temp_dir = base_dir / "temp"
    # Link do chromedriver (substitua pelo link do Onedrive se necessário)
    # onedrive_link = "https://onedrive.live.com/download?cid=EF1D4DBF39C8FC44&resid=EF1D4DBF39C8FC44!1133396&authkey=!AD2dvS2GMjWFYsE"  # Certifique-se de que o link está correto para a versão desejada
    #servidor_link = "http://seu_servidor.com/pasta/chromedriver.zip"
    caminho_zip_local = "C:\\Users\\erics\\OneDrive\\TI\\PROJETOS\\Habitacao\\chromedriver.zip"

    # Verifica se o chromedriver existe na pasta destino e baixa, se necessário
    if not chromedriver_path.exists():
        print("Chromedriver não encontrado localmente. Baixando e extraindo da pasta origem ...")
        #extrair_chromedriver(caminho_zip_local, str(chromedriver_path), str(temp_dir))
        # baixar_e_extrair_chromedriver(onedrive_link, str(chromedriver_path), str(temp_dir))
        extrair_chromedriver_local(caminho_zip_local, str(chromedriver_path), str(temp_dir))
    else:
        print("Chromedriver já está configurado.")

    # Inicializa o webdriver
    inicializar_navegador(str(chromedriver_path))

    # Remove o diretório temporário após a execução
    shutil.rmtree(str(temp_dir), ignore_errors=True)

except Exception as e:
    # Mostra o alerta com a mensagem de erro
    exibir_alerta("Ocorreu um erro inesperado: {}".format(str(e)))
    sys.exit(1)  # Código de erro 1 para erro inesperado

