from pathlib import Path
import os

# Obtém o diretório do perfil do usuário atual
user_profile = os.getenv("USERPROFILE")

# Define o caminho do chromedriver com base no diretório do perfil do usuário
CHROMEDRIVER_PATH = Path(f"{user_profile}/AppData/Local/chromedriver/chromedriver.exe")
print(CHROMEDRIVER_PATH)