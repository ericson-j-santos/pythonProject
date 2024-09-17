# pip install selenium
from selenium import webdriver
# pip install webdriver-manager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# Se ao usar a estrutura abaixo com service e aparecer erro na execucção com relação ao service, trata-se da estrutura abaixo que é referente a nova atualizacao do selenium e para isso deve-se fazer: pip install --upgrade selenium OU pip install selenium==4.1.0
servico = Service(ChromeDriverManager().install())
# servico = Service(r'C:\Users\erics\Downloads\120.0.6099.109\chrome-win64 120.0.6099.109\chrome-win64\chrome.exe')

navegador = webdriver.Chrome(service=servico)

navegador.get("https://vuetifyjs.com/en/components/date-pickers/#usage")






# navegador.get('https://www.google.com/')
# navegador.maximize_window()
