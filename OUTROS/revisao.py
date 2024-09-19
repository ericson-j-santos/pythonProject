import openpyxl

# Função para carregar os resultados do arquivo Excel
def carregar_resultados_do_excel(arquivo_excel):
    try:
        workbook = openpyxl.load_workbook(arquivo_excel)
        sheet = workbook.active
        resultados = []
        # Começamos a partir da segunda linha para pular o cabeçalho
        for row in sheet.iter_rows(min_row=2, values_only=True):
            arquivo, link = row
            resultados.append({"arquivo": arquivo, "link": link})
        return resultados
    except FileNotFoundError:
        print(f"Arquivo {arquivo_excel} não encontrado.")
        return []

# Exibir os resultados para revisão
arquivo_resultado_excel = 'resultados_pesquisa_google.xlsx'
resultados_pesquisa = carregar_resultados_do_excel(arquivo_resultado_excel)

# Exibindo uma amostra dos resultados
print("Exibindo uma amostra dos resultados:")
for resultado in resultados_pesquisa[:10]:  # Exibe os primeiros 10 resultados
    print(f"Arquivo: {resultado['arquivo']}\nLink: {resultado['link']}\n")
