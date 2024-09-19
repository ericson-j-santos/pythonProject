import matplotlib.pyplot as plt

# Função para gerar gráficos de validação
def gerar_relatorio(resultados_validados):
    # Contagem de resultados válidos e inválidos
    validos = sum(1 for resultado in resultados_validados if resultado["valido"])
    invalidos = len(resultados_validados) - validos

    # Criando um gráfico de pizza
    labels = 'Válidos', 'Inválidos'
    sizes = [validos, invalidos]
    colors = ['#4CAF50', '#F44336']
    explode = (0.1, 0)  # Destaque para válidos

    plt.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')  # Assegura que o gráfico seja um círculo
    plt.title("Distribuição de Resultados Válidos vs Inválidos")
    plt.show()

# Gerando o relatório
gerar_relatorio(resultados_validados)
