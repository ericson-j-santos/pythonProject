# código original
import os
import time
import shutil
import sys
import datetime
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PDL.src.backup.chromedriver_configura import configurar_chromedriver, inicializar_navegador, exibir_alerta, show_message
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

        # Espera explícita até que o campo de usuário esteja presente
        WebDriverWait(navegador,35).until(
            EC.presence_of_element_located((By.XPATH, login_xpath))
        )

        # Inserir o nome de usuário
        navegador.find_element(By.XPATH, login_xpath).send_keys(user)
        navegador.find_element(By.ID, login_button_id).click()

        # Solicitar ao usuário que insira o código de verificação manualmente
        show_message()

        # Espera explícita de até 30 segundos para o elemento ID Carteira para selecionar e carteira esteja presente
        WebDriverWait(navegador, 30).until(
            EC.presence_of_element_located((By.ID, '//*[@id="Carteira"]'))
        )
        time.sleep(5)

        navegador.find_element(By.XPATH, '//*[@id="Carteira"]').click()
        time.sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="Carteira"]/option[7]').click()

        time.sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="Menu"]').send_keys("PDL")

        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="entrarMenu"]/i'))
        )

        navegador.find_element(By.XPATH, '//*[@id="entrarMenu"]/i').click()

        # Verificar se o elemento de menu está presente para confirmar o login
        print("Login realizado com sucesso!")
    except Exception as e:
        print(f"Erro ao logar e navegar: {e}")
        return False
    return True


# Função para processar contratos e preencher dados de liquidação na planilha .xlsm
def processar_contratos_xlsm(navegador, planilha):
    try:
        # Conectar ao Excel e abrir a planilha
        excel = win32.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(planilha)
        ws = workbook.Sheets("DADOS")

        row = 2  # Começa na linha 2, assumindo que a linha 1 tem os cabeçalhos

        while True:
            # 1. Ler o número do contrato da planilha (coluna B)
            nr_contrato_planilha = ws.Cells(row,2).value  # Coluna B (index 2)

            # Verificar se a célula do contrato está vazia, se sim, parar o processamento
            if nr_contrato_planilha is None:
                print(f"Célula de contrato vazia na linha {row}, parando o processamento.")
                break  # Se estiver vazia, para o loop
            print(f"Processando o contrato: {nr_contrato_planilha} na linha {row}")

            # 2. Extrair as datas das colunas (H, J, L)
            colunas_datas = [8, 10, 12]  # Colunas H, J, L onde estão as datas
            datas_planilha = []

            for col in colunas_datas:
                data = ws.Cells(row, col).value
                # Verificar se a data é datetime e formatar para o formato dd/mm/yyyy
                if isinstance(data, datetime.datetime):
                    data = data.strftime("%d/%m/%Y")
                datas_planilha.append(data)  # Adicionar cada data à lista
            print(f"Datas extraídas: {datas_planilha}")

            # 3. Preencher o número do contrato no formulário na web
            WebDriverWait(navegador,30).until(
                EC.presence_of_element_located((By.ID, "numero-contrato"))
            )
            campo_contrato = navegador.find_element(By.ID, "numero-contrato")
            campo_contrato.clear()
            campo_contrato.send_keys(str(nr_contrato_planilha))

            # 4. Selecionar o tipo de liquidação (por exemplo, "Liquidação Por Portabilidade")
            WebDriverWait(navegador,30).until(
                EC.presence_of_element_located((By.NAME, "CodigoTipoLiquidacao"))
            )
            tipo_liq_web = navegador.find_element(By.NAME, "CodigoTipoLiquidacao")
            option_tipo_liq_web = tipo_liq_web.find_element(By.XPATH, "//option[text()='Liquidação Por Portabilidade']")
            option_tipo_liq_web.click()

            # 5. Preencher as datas de liquidação da planilha no formulário da web
            for i, data in enumerate(datas_planilha):
                if data: # Verifica se há uma data presente
                    print(f"Processando Data {i+1} para o contrato {nr_contrato_planilha}")
                    WebDriverWait(navegador,30).until(
                        EC.presence_of_element_located((By.NAME, "DataReferencia"))
                    )
                    dt_liq_web = navegador.find_element(By.NAME, "DataReferencia")
                    dt_liq_web.clear()
                    dt_liq_web.send_keys(data)

                    #  Clicar no botão de consulta para processar a liquidação e aguardar o saldo ser gerado
                    navegador.find_element(By.ID, "opcCON").click()

                    #  Esperar até que o saldo gerado apareça no campo específico e capturar o valor
                    WebDriverWait(navegador,30).until(
                        EC.presence_of_element_located((By.ID, "valorDividaLiquidacao"))
                    )
                    valorDivida = navegador.find_element(By.ID, "valorDividaLiquidacao").get_attribute("value")
                    print(f"Valor da dívida para Data {i+1}: {valorDivida}")

                    #  Preencher o saldo correspondente na planilha nas colunas de saldo (colunas I, K, M)
                    # coluna_saldo = 9 + (i *2)
                    # ws.Cells(row, coluna_saldo).Value = valorDivida  # Colunas I, K, M
                    ws.Cells(row, 9 + (i *2)).Value = valorDivida  # Colunas I, K, M
                    # print(f"Valor {valorDivida} preenchido na célula ({row}, {coluna_saldo})")

                    # Espera antes de processar a próxima data para garantir que a página foi atualizada
                    time.sleep(2)  # Ajuste conforme necessário

            # 6. Preencher o "Saldo do Dia" com a data atual
            print(f"Processando Saldo do Dia para o contrato {nr_contrato_planilha}")
            data_hoje = datetime.datetime.now().strftime("%d/%m/%Y")
            WebDriverWait(navegador, 30).until(
                EC.presence_of_element_located((By.NAME, "DataReferencia"))  # Usando o campo genérico de data
            )
            dt_liq_web = navegador.find_element(By.NAME, "DataReferencia")
            dt_liq_web.clear()
            dt_liq_web.send_keys(data_hoje)

            # Clicar no botão de consulta novamente para obter o saldo do dia
            navegador.find_element(By.ID, "opcCON").click()

            time.sleep(5)

            # Capturar o saldo do dia
            WebDriverWait(navegador, 30).until(
                EC.presence_of_element_located((By.ID, "valorDividaLiquidacao"))
            )
            saldo_do_dia = navegador.find_element(By.ID, "valorDividaLiquidacao").get_attribute("value")
            print(f"Saldo do dia: {saldo_do_dia}")

            # Preencher o saldo do dia na coluna "Saldo do Dia" (Coluna R)
            ws.Cells(row, 18).Value = saldo_do_dia  # Coluna R

            # 7. Atualizar a página (refresh) para garantir que os dados sejam limpos e processar o próximo contrato
            navegador.refresh()

            # Continuar para a próxima linha da planilha
            row += 1

        # 8. Salvar a planilha com os valores preenchidos
        workbook.Save()
        print("Processamento de contratos concluído!")
        # workbook.Close()

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
