import os
from selenium import webdriver

try:
    # Obtém o nome do usuário atual
    nome_usuario = os.getlogin()

    # Caminho para diretório 'AppData\Local' do usuário
    base_dir = os.path.join("C:\\Users", nome_usuario, "AppData", "Local")

    # Caminho para a pasta onde o 'chromedriver.exe' foi configurado
    chromedriver_path = os.path.join(base_dir, "chromedriver", "chromedriver.exe")

    # Inicializa o webdriver
    navegador = webdriver.Chrome(executable_path=chromedriver_path)

    # Ação de raspagem de dados
    navegador.get("https://google.com")

    # Exemplo de raspagem de dados
    dados = navegador.find_element_by_tag_name("body").text
    print("Dados raspados", dados)

    navegador.quit()

except Exception as e:
    print(f"Erro durante a raspagem de dados: {str(e)}")
