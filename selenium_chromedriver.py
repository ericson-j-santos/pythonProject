from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#navegador = webdriver.Chrome(r'C:\Users\erics\Downloads\120.0.6099.109\chromedriver-win64 120.0.6099.109\chromedriver-win64\chromedriver')
#navegador.get("https://demo.automationtesting.in/Datepicker.html")

def ep27_execute_javascript():
    # Caminho manual para o ChromeDriver, já configurado e funcionando
    driver_path = r'C:\Users\erics\Downloads\120.0.6099.109\chromedriver-win64 120.0.6099.109\chromedriver-win64\chromedriver'

    # Caminho para o binário do Google Chrome versão 120
    chrome_binary_path = r'C:\Users\erics\Downloads\120.0.6099.109\chromedriver-win64 120.0.6099.109\chromedriver-win64\chrome.exe'

    # Configurações do Chrome
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.binary_location = chrome_binary_path  # Define o caminho para o binário do Chrome versão 120

    # Inicia o navegador Chrome com o ChromeDriver especificado manualmente
    driver = webdriver.Chrome(executable_path=driver_path, chrome_options=options)
    driver.get("https://demo.automationtesting.in/Datepicker.html")

    # Insere o valor de A3 (por exemplo, "03/07/2022") no campo de data com ID 'datepicker2'
    data_a3 = "03/07/2022"  # Este valor seria extraído de uma célula de uma planilha no VBA
    driver.find_element_by_id("datepicker2").send_keys(data_a3)

    # Clica no elemento identificado pelo XPath
    driver.find_element_by_xpath("//label[text()='DatePicker Enabled']").click()

    # Remove o atributo 'readonly' do campo de data com ID 'datepicker1' usando JavaScript
    driver.execute_script("arguments[0].removeAttribute('readonly')", driver.find_element_by_id("datepicker1"))

    # Insere uma data no campo de data com ID 'datepicker1'
    driver.find_element_by_id("datepicker1").send_keys("03/07/2022")
    driver.find_element_by_id("datepicker1").send_keys(Keys.TAB)

    # Adiciona novamente o atributo 'readonly' ao campo de data com ID 'datepicker1' usando JavaScript
    driver.execute_script("arguments[0].setAttribute('readonly', 'true')", driver.find_element_by_id("datepicker1"))

    # Fecha o navegador
    driver.quit()

# Executa a função
ep27_execute_javascript()

#-----------------------------------------------------------------------------------------------------------------------
# Código usado em conjunto no VBA
# Sub RunPythonScript()
#
# Dim objShell As Object
# Dim PythonExePath, PythonScriptPath As String
#
#     Set objShell = VBA.CreateObject("Wscript.Shell")
#     'Comando no CMD para saber o caminho do python: where python
#     PythonExePath = """C:\Users\erics\anaconda3\python.exe"""
#     PythonScriptPath = """C:\Users\erics\PycharmProjects\pythonProject\selenium_chromedriver.py"""
#
#     objShell.Run
#     PythonExePath & PythonScriptPath
#
# End Sub