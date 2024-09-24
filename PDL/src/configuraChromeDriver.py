import os
import sys
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from tkinter import messagebox, Tk
from pathlib import Path
#import urllib.request
import zipfile

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
#     try:
#         # Cria o diretório temporário
#         if not os.path.exists(temp_dir):
#             os.makedirs(temp_dir)
#
#         # Caminho do arquivo zip temporário
#         zip_path = os.path.join(temp_dir, "chromedriver.zip")
#
#         # Baixa o arquivo zip da pasta origem
#         urllib.request.urlretrieve(url, zip_path)  # Alterado para urllib compatível com Python 2.7
#         print("Arquivo baixado para: {}".format(zip_path))
#
#         # Extrai o arquivo zip
#         with zipfile.ZipFile(zip_path, "r") as zip_ref:
#             zip_ref.extractall(temp_dir)
#             print("Arquivo extraído para {}".format(temp_dir))
#
#         # Procura pelo chromedriver.exe dentro de uma pasta extraída
#         for root, dirs, files in os.walk(temp_dir):
#             if "chromedriver.exe" in files:
#                 chromedriver_exe = os.path.join(root, "chromedriver.exe")
#                 shutil.move(chromedriver_exe, destino)
#                 print("Chromedriver copiado para: {}".format(destino))
#                 break
#         else:
#             exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
#             sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip
#
#         # Remove o arquivo zip após extração
#         os.remove(zip_path)
#
#     except Exception as e:
#         exibir_alerta("Falha ao baixar e extrair o chromedriver: {}".format(str(e)))
#         sys.exit(2)  # Código de erro 2 para falha ao baixar e extrair arquivo



# Função para extrair o chromedriver de um arquivo zip local


def extrair_chromedriver(caminho_zip, destino, temp_dir):
    try:
        # Cria o diretório temporário
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # Extrai o arquivo zip
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            print("Arquivo extraído para {}".format(temp_dir))

        # Procura pelo chromedriver.exe dentro de uma pasta extraída
        for root, dirs, files in os.walk(temp_dir):
            if "chromedriver.exe" in files:
                chromedriver_exe = os.path.join(root, "chromedriver.exe")
                shutil.move(chromedriver_exe, destino)
                print("Chromedriver copiado para: {}".format(destino))
                break
        else:
            exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
            sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip

    except Exception as e:
        exibir_alerta("Falha ao extrair o chromedriver: {}".format(str(e)))
        sys.exit(2)  # Código de erro 2 para falha ao extrair arquivo

# Função para inicializar o webdriver
def inicializar_navegador(chromedriver_path):
    try:

        # Debugging: Verificando o ambiente e o caminho do chromedriver
        print("Python Executável:", sys.executable)
        print("Versão do Python:", sys.version)
        print("Caminho do chromedriver:", chromedriver_path)

        #Configurar as opções do Chrome
        chrome_options = Options()
        #chrome_options.add_argument("--headless") # Rodar o navegador sem interface gráfica (opcional)
        # E adiciona esse trecho na linha debaixo, conforme exemplo abaixo, options=chrome_options ...
        #navegador = webdriver.Chrome(executable_path=chromedriver_path, options=chrome_options)

        #---------------------------------------------------
        # Criar o objeto Service com o caminho do chromedriver
        service = Service(executable_path=chromedriver_path)
        # Inicializa o navegador com o driver configurado
        #navegador = webdriver.Chrome(service=service, options=chrome_options)
        # ---------------------------------------------------

        # ---------------------------------------------------
        # Inicializa o webdriver diretamente com o caminho do chromedriver
        #navegador = webdriver.Chrome(executable_path=chromedriver_path, options=chrome_options) - nâo deu certo
        navegador = webdriver.Chrome(service=service, options=chrome_options)  # Ajuste aqui, passando o caminho como argumento posicional

        # ---------------------------------------------------

        # Define timeout explicitamente (opcional, se necessário)
        navegador.set_page_load_timeout(30)  # Define timeout de 30 segundos para carregamento da página

        navegador.get("https://www.google.com.br")
        print("Navegador foi aberto com sucesso")
        navegador.quit()
        exibir_sucesso("Navegador foi aberto com sucesso.")
    except Exception as e:
        exibir_alerta("Ocorreu um erro ao abrir o navegador: {}".format(str(e)))
        sys.exit(4)  # Código de erro 4 para falha ao inicializar navegador

try:
    # Caminhos de configuração
    # Obtém o nome do usuário atual
    nome_usuario = os.getlogin()

    # Caminho para o diretório 'AppData\Local' do usuário
    base_dir = Path("C:/Users/{}/AppData/Local/chromedriver".format(nome_usuario))

    # Caminho para a pasta onde o 'chromedriver.exe' será colocado
    chromedriver_path = base_dir / "chromedriver.exe"
    temp_dir = base_dir / "temp"

    # Link do chromedriver (substitua pelo link do Onedrive se necessário)
    #onedrive_link = "https://onedrive.live.com/download?cid=EF1D4DBF39C8FC44&resid=EF1D4DBF39C8FC44!1133396&authkey=!AD2dvS2GMjWFYsE"  # Certifique-se de que o link está correto para a versão desejada
    #servidor_link = "http://seu_servidor.com/pasta/chromedriver.zip"
    caminho_zip_local = "C:\\Users\\erics\\OneDrive\\TI\\PROJETOS\\Habitacao\\chromedriver.zip"

    # Verifica se o chromedriver existe na pasta destino e baixa, se necessário
    if not chromedriver_path.exists():
        print("Chromedriver não encontrado localmente. Baixando e extraindo da pasta origem ...")
        extrair_chromedriver(caminho_zip_local, str(chromedriver_path), str(temp_dir))
        #baixar_e_extrair_chromedriver(onedrive_link, str(chromedriver_path), str(temp_dir))
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








# -*- coding: utf-8 -*-
# Codigo Ok - usando link do chromedriver no onedriver;
# Pacotes e versões = Python 3.9; selenium 3.141.0
# pip install -r requirements.txt

# import os
# import sys
# import shutil
# from selenium import webdriver
# #from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from tkinter import messagebox, Tk
# from pathlib import Path
# import urllib.request
# import zipfile
#
# # Função para exibir uma mensagem de erro em um alerta
# def exibir_alerta(mensagem):
#     root = Tk()
#     root.withdraw()  # Esconde a janela principal do tkinter
#     messagebox.showerror("Erro", mensagem)
#     root.destroy()  # Fecha a janela tkinter após o alerta
#
# # Função para exibir uma mensagem de sucesso em um alerta
# def exibir_sucesso(mensagem):
#     root = Tk()
#     root.withdraw()  # Esconde a janela principal do tkinter
#     messagebox.showinfo("Sucesso", mensagem)
#     root.destroy()  # Fecha a janela tkinter após o alerta
#
# # Função para baixar e extrair o chromedriver de um arquivo zip do local de origem
# # def baixar_e_extrair_chromedriver(url, destino, temp_dir):
# #     try:
# #         # Cria o diretório temporário
# #         if not os.path.exists(temp_dir):
# #             os.makedirs(temp_dir)
# #
# #         # Caminho do arquivo zip temporário
# #         zip_path = os.path.join(temp_dir, "chromedriver.zip")
# #
# #         # Baixa o arquivo zip da pasta origem
# #         urllib.request.urlretrieve(url, zip_path)  # Alterado para urllib compatível com Python 2.7
# #         print("Arquivo baixado para: {}".format(zip_path))
# #
# #         # Extrai o arquivo zip
# #         with zipfile.ZipFile(zip_path, "r") as zip_ref:
# #             zip_ref.extractall(temp_dir)
# #             print("Arquivo extraído para {}".format(temp_dir))
# #
# #         # Procura pelo chromedriver.exe dentro de uma pasta extraída
# #         for root, dirs, files in os.walk(temp_dir):
# #             if "chromedriver.exe" in files:
# #                 chromedriver_exe = os.path.join(root, "chromedriver.exe")
# #                 shutil.move(chromedriver_exe, destino)
# #                 print("Chromedriver copiado para: {}".format(destino))
# #                 break
# #         else:
# #             exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
# #             sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip
# #
# #         # Remove o arquivo zip após extração
# #         os.remove(zip_path)
# #
# #     except Exception as e:
# #         exibir_alerta("Falha ao baixar e extrair o chromedriver: {}".format(str(e)))
# #         sys.exit(2)  # Código de erro 2 para falha ao baixar e extrair arquivo
#
#
#
# # Função para extrair o chromedriver de um arquivo zip local
#
#
# def extrair_chromedriver(caminho_zip, destino, temp_dir):
#     try:
#         # Cria o diretório temporário
#         if not os.path.exists(temp_dir):
#             os.makedirs(temp_dir)
#
#         # Extrai o arquivo zip
#         with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
#             zip_ref.extractall(temp_dir)
#             print("Arquivo extraído para {}".format(temp_dir))
#
#         # Procura pelo chromedriver.exe dentro de uma pasta extraída
#         for root, dirs, files in os.walk(temp_dir):
#             if "chromedriver.exe" in files:
#                 chromedriver_exe = os.path.join(root, "chromedriver.exe")
#                 shutil.move(chromedriver_exe, destino)
#                 print("Chromedriver copiado para: {}".format(destino))
#                 break
#         else:
#             exibir_alerta("Arquivo chromedriver.exe não encontrado no zip.")
#             sys.exit(3)  # Código de erro 3 para o chromedriver não encontrado no zip
#
#     except Exception as e:
#         exibir_alerta("Falha ao extrair o chromedriver: {}".format(str(e)))
#         sys.exit(2)  # Código de erro 2 para falha ao extrair arquivo
#
# # Função para inicializar o webdriver
# def inicializar_navegador(chromedriver_path):
#     try:
#
#         # Debugging: Verificando o ambiente e o caminho do chromedriver
#         print("Python Executável:", sys.executable)
#         print("Versão do Python:", sys.version)
#         print("Caminho do chromedriver:", chromedriver_path)
#
#         #Configurar as opções do Chrome
#         chrome_options = Options()
#         #chrome_options.add_argument("--headless") # Rodar o navegador sem interface gráfica (opcional)
#         # E adiciona esse trecho na linha debaixo, conforme exemplo abaixo, options=chrome_options ...
#         #navegador = webdriver.Chrome(executable_path=chromedriver_path, options=chrome_options)
#
#         #---------------------------------------------------
#         # Criar o objeto Service com o caminho do chromedriver
#         #service = Service(executable_path=chromedriver_path)
#         # Inicializa o navegador com o driver configurado
#         #navegador = webdriver.Chrome(service=service, options=chrome_options)
#         # ---------------------------------------------------
#
#         # ---------------------------------------------------
#         # Inicializa o webdriver diretamente com o caminho do chromedriver
#         #navegador = webdriver.Chrome(executable_path=chromedriver_path, options=chrome_options) - nâo deu certo
#         navegador = webdriver.Chrome(chromedriver_path, chrome_options)  # Ajuste aqui, passando o caminho como argumento posicional
#
#         # ---------------------------------------------------
#
#         # Define timeout explicitamente (opcional, se necessário)
#         navegador.set_page_load_timeout(30)  # Define timeout de 30 segundos para carregamento da página
#
#         navegador.get("https://www.google.com.br")
#         print("Navegador foi aberto com sucesso")
#         navegador.quit()
#         exibir_sucesso("Navegador foi aberto com sucesso.")
#     except Exception as e:
#         exibir_alerta("Ocorreu um erro ao abrir o navegador: {}".format(str(e)))
#         sys.exit(4)  # Código de erro 4 para falha ao inicializar navegador
#
# try:
#     # Caminhos de configuração
#     # Obtém o nome do usuário atual
#     nome_usuario = os.getlogin()
#
#     # Caminho para o diretório 'AppData\Local' do usuário
#     base_dir = Path("C:/Users/{}/AppData/Local/chromedriver".format(nome_usuario))
#
#     # Caminho para a pasta onde o 'chromedriver.exe' será colocado
#     chromedriver_path = base_dir / "chromedriver.exe"
#     temp_dir = base_dir / "temp"
#
#     # Link do chromedriver (substitua pelo link do Onedrive se necessário)
#     #onedrive_link = "https://onedrive.live.com/download?cid=EF1D4DBF39C8FC44&resid=EF1D4DBF39C8FC44!1133396&authkey=!AD2dvS2GMjWFYsE"  # Certifique-se de que o link está correto para a versão desejada
#     #servidor_link = "http://seu_servidor.com/pasta/chromedriver.zip"
#     caminho_zip_local = "C:\\Users\\erics\\OneDrive\\TI\\PROJETOS\\Habitacao\\chromedriver.zip"
#
#     # Verifica se o chromedriver existe na pasta destino e baixa, se necessário
#     if not chromedriver_path.exists():
#         print("Chromedriver não encontrado localmente. Baixando e extraindo da pasta origem ...")
#         extrair_chromedriver(caminho_zip_local, str(chromedriver_path), str(temp_dir))
#         #baixar_e_extrair_chromedriver(onedrive_link, str(chromedriver_path), str(temp_dir))
#     else:
#         print("Chromedriver já está configurado.")
#
#     # Inicializa o webdriver
#     inicializar_navegador(str(chromedriver_path))
#
#     # Remove o diretório temporário após a execução
#     shutil.rmtree(str(temp_dir), ignore_errors=True)
#
# except Exception as e:
#     # Mostra o alerta com a mensagem de erro
#     exibir_alerta("Ocorreu um erro inesperado: {}".format(str(e)))
#     sys.exit(1)  # Código de erro 1 para erro inesperado





