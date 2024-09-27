from docx.opc.oxml import oxml_parser

from consultaPDL import login_e_navegar
from chromedriver_configura import inicializar_navegador, exibir_alerta
from pathlib import Path
import os

CHROMEDRIVER_PATH = Path("C:/Users/erics/AppData/Local/chromedriver/chromedriver.exe")
URL_LOGIN = "https://accounts.google.com/v3/signin/identifier?hl=pt-BR&ifkv=ARpgrqeQYawdtX3RE9Wg1hB4W4BAdnFdJfZ9fXHgmA-2814xb_f9ghCyRXxtCKpdjq-X5cDWhfxc&ddm=0&flowName=GlifWebSignIn&flowEntry=ServiceLogin&continue=https%3A%2F%2Faccounts.google.com%2FManageAccount%3Fnc%3D1"
LOGIN_XPATH = '//*[@id="identifierId"]'
LOGIN_BUTTON_ID = "identifierNext"

# Inicializar o navegador
navegador = inicializar_navegador(str(CHROMEDRIVER_PATH))

# Obter nome de usuário do sistema
usuario = os.getlogin()
print(f"Nome do usuário: {usuario}")

# Testar a função de login
resultado = login_e_navegar(navegador, URL_LOGIN, usuario, LOGIN_XPATH, LOGIN_BUTTON_ID)
if resultado:
    print("Teste de login e navegação bem-sucedido.")
else:
    print("Falha no teste de login e navegação.")

# Fechar o navegador
navegador.quit()