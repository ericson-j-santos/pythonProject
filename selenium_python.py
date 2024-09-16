from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

servico = Service(ChromeDriverManager().install())
# servico = Service(r'C:\Users\erics\Downloads\120.0.6099.109\chrome-win64 120.0.6099.109\chrome-win64\chrome.exe')

navegador = webdriver.Chrome(service=servico)






# navegador.get('https://www.google.com/')
# navegador.maximize_window()
