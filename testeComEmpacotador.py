import os
import shutil
from selenium import webdriver
from tkinter import messagebox, Tk

# Função para exibir uma mensagem de erro em um alerta
def exibir_alerta(mensagem):
    root = Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    messagebox.showerror("Erro", mensagem)
    root.destroy()  # Fecha a janela tkinter após o alerta

# Função para copiar o chromedriver para a pasta destino se não existir
def copiar_chromedriver(origem, destino):
    if not os.path.exists(destino):
        try:
            # Cria a pasta de destino se não existir
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            # Copia o chromedriver do caminho de origem para o destino
            shutil.copyfile(origem, destino)
            print(f"Chromedriver copiado para : {destino}")
        except Exception as e:
            exibir_alerta(f"Falha ao copia o chromedriver: {str(e)}")


try:
    # Obtém o nome do usuário atual
    nome_usuario = os.getlogin()

    # Caminho para o diretório 'AppData\Local' do usuário
    base_dir = os.path.join("C:\\Users", nome_usuario, "AppData", "Local")

    # Caminho para a pasta onde o 'chromedriver.exe' será colocado
    chromedriver_path = os.path.join(base_dir, "chromedriver", "chromedriver.exe")

    # Caminho do chromedriver embutido no projeto (diretório 'resources')
    chromedriver_embutido = os.path.join(os.getcwd(), "resources", "chromedriver.exe")

    # Se o chromedriver não estiver na pasta, copia-o do projeto
    if not os.path.exists(chromedriver_path):
        copiar_chromedriver(chromedriver_embutido, chromedriver_path)

    # Inicializa o WebDriver
    navegador = webdriver.Chrome(executable_path=chromedriver_path)

    # Ação do WebDriver (exemplo: abrir uma página)
    navegador.get("https://www.google.com")
    print("Navegador foi aberto com sucesso.")
    navegador.quit()

except Exception as e:
    # Mostra o alerta com a mensagem de erro
    exibir_alerta(f"Ocorreu um erro: {str(e)}")
