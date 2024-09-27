# código original
import os
import time
import shutil
import sys
from pathlib import Path
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from chromedriver_configura import configurar_chromedriver, inicializar_navegador, exibir_alerta
import win32com.client as win32 # Biblioteca para interagir com o excel

# Função para obter a planilha aberta no Excel
def obter_planilha_aberta():
    try:
        # Conectar ao Excel em execução
        excel = win32.Dispatch("Excel.Application")
        # Seleciona a planilha ativa
        workbook = excel.ActiveWorkbook
        # Retorna o caminho completo da planilha ativa
        planilha_aberta = workbook.FullName
        return planilha_aberta
    except Exception as e:
        exibir_alerta(f"Erro ao acessar a planilha aberta {str(e)}")
        return None

# Função para fazer login no portal e navegar para a tela de PDL
def login_e_navegar(navegador, url, user, login_xpath, login_button_id):
    try:
        # Navegar até a página de login
        navegador.get(url)
        time.sleep(2)  # Aguardar a página carregar

        # Inserir o nome de usuário
        navegador.find_element(By.XPATH, login_xpath).send_keys(user)
        navegador.find_element(By.ID, login_button_id).click()

        # Solicitar ao usuário que insira o código de verificação manualmente
        input("Insira o código de verificação no navegador e pressione Enter para continuar...")

        # Verificar se o elemento de menu está presente para confirmar o login
        navegador.find_element(By.ID, "Menu")
        print("Login realizado com sucesso!")
    except Exception as e:
        print(f"Erro durante o login e navegação: {e}")
        return False
    return True


# Função para processar contratos e preencher dados de liquidação na planilha .xlsm
def processar_contratos_xlsm(navegador, planilha):
    try:
        # Conectar ao Excel e abrir a planilha
        excel = win32.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(planilha)
        ws = workbook.Sheets("Dados")

        # Iterar sobre cada linha da planilha, começando da segunda linha
        for row in range(2, ws.UsedRange.Rows.Count + 1):
            nr_contrato_planilha = ws.Cells(row,2).value  # Coluna B (index 1)

            # Preencher o campo de contrato no navegador
            campo_contrato = navegador.find_element(By.ID, "numero-contrato")
            campo_contrato.clear()
            campo_contrato.send_keys(nr_contrato_planilha)
            time.sleep(2)  # Aguardar o preenchimento

            # Selecionar o tipo de liquidação
            tipo_liq_web = navegador.find_element(By.NAME, "CodigoTipoLiquidacao")
            option_tipo_liq_web = tipo_liq_web.find_element(By.XPATH, "//option[text()='Liquidação Por Portabilidade']")
            option_tipo_liq_web.click()
            time.sleep(2)  # Aguardar seleção

            # Processar cada data de liquidação na planilha e preencher o saldo correspondente
            for i in range(7, 13, 2):  # Colunas H (index 7), J (index 9), L (index 11)
                data_liq_planilha = ws.Cells(row, i).value
                coluna_saldo = i + 1 # Coluna correspondente ao saldo (I, K, M)

                if data_liq_planilha is None:
                    continue

                # Atualizar a página para garantir a atualização de dados
                navegador.refresh()
                time.sleep(2)

                # Preencher a data de liquidação
                dt_liq_web = navegador.find_element(By.NAME, "DataReferencia")
                dt_liq_web.clear()
                dt_liq_web.send_keys(data_liq_planilha)
                time.sleep(2)

                # Clicar no botão de consulta
                navegador.find_element(By.ID, "opcCON").click()
                time.sleep(2)

                # Aguardar até que o elemento valorDividaLiquidacao apareça e capturar o valor
                try:
                    valorDivida = navegador.find_element(By.ID, "valorDividaLiquidacao").text
                except Exception:
                    valorDivida = "N/A" # Em caso de erro, atribuir valor não disponível
                print(f"Valor da dívida: {valorDivida}")

                # Preencher o valor na planilha na coluna correspondente ao saldo
                ws.cell(row, coluna_saldo).value = valorDivida

        # Salvar a planilha
        workbook.Save()
        print("Processamento de contratos concluído!")
        workbook.Close()
    except Exception as e:
        print(f"Erro ao processar contratos: {e}")


# Função principal para executar o script
def executar_script(chromedriver_path, planilha, url, login_xpath, login_button_id):
    try:
        # Configura o chromedriver
        caminho_zip_local = "resources/chromedriver.zip"
        temp_dir = Path(chromedriver_path).parent / "temp"
        configurar_chromedriver(caminho_zip_local, chromedriver_path, temp_dir)

        # Inicializa o navegador
        navegador = inicializar_navegador(str(chromedriver_path))
        if not navegador:
            return

        # Pega o nome do usuário do sistema
        usuario = os.getlogin()

        # Faz o login e navega
        if not login_e_navegar(navegador, url, usuario, login_xpath, login_button_id):
            return

        # Processa os contratos
        processar_contratos_xlsm(navegador, planilha)

        # Fecha o navegador
        navegador.quit()

        # Remove o diretório temporário após a execução
        shutil.rmtree(str(temp_dir), ignore_errors=True)
    except Exception as e:
        exibir_alerta(f"Ocorreu um erro inesperado: {str(e)}")
        sys.exit(1)
    # processar_contratos(navegador, planilha)

# Definição dos caminhos e variáveis
CHROMEDRIVER_PATH = Path("C:/Users/{}/AppData/Local/chromedriver/chromedriver.exe".format(os.getlogin()))

PLANILHA = obter_planilha_aberta()

URL_LOGIN = "https://accounts.google.com/v3/signin/identifier?hl=pt-BR&ifkv=ARpgrqeQYawdtX3RE9Wg1hB4W4BAdnFdJfZ9fXHgmA-2814xb_f9ghCyRXxtCKpdjq-X5cDWhfxc&ddm=0&flowName=GlifWebSignIn&flowEntry=ServiceLogin&continue=https%3A%2F%2Faccounts.google.com%2FManageAccount%3Fnc%3D1"
LOGIN_XPATH = '//*[@id="identifierId"]'
LOGIN_BUTTON_ID = "identifierNext"

if PLANILHA: # Verifica se a planilha foi encontrada
    if __name__ == "__main__":
        executar_script(CHROMEDRIVER_PATH, PLANILHA, URL_LOGIN, LOGIN_XPATH, LOGIN_BUTTON_ID)
else:
    exibir_alerta("Nenhuma planilha aberta foi encontrada. Por favor, abra uma planilha no Excel.")
