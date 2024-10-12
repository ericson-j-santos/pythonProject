import matplotlib.pyplot as plt
import pandas as pd
import pickle
from collections import Counter
from wordcloud import WordCloud

# Carregar os resultados validados do arquivo
with open('resultados_validados.pkl', 'rb') as f:
    resultados_validados = pickle.load(f)

# Resumo estatístico
total_registros = len(resultados_validados)
total_validos = sum(1 for resultado in resultados_validados if resultado["valido"] == True)
total_invalidos = total_registros - total_validos
erro_404 = sum(1 for resultado in resultados_validados if "404" in resultado.get("erro", ""))
erro_403 = sum(1 for resultado in resultados_validados if "403" in resultado.get("erro", ""))
erro_429 = sum(1 for resultado in resultados_validados if "429" in resultado.get("erro", ""))
erro_timeout = sum(1 for resultado in resultados_validados if "timeout" in resultado.get("erro", ""))

# Exibindo resumo estatístico
print(f"Total de Registros Processados: {total_registros}")
print(f"Total de Registros Válidos: {total_validos} ({(total_validos / total_registros) * 100:.2f}%)")
print(f"Total de Registros Inválidos: {total_invalidos} ({(total_invalidos / total_registros) * 100:.2f}%)")
print(f"Total de Erros 404: {erro_404}")
print(f"Total de Erros 403: {erro_403}")
print(f"Total de Erros 429: {erro_429}")
print(f"Total de Erros de Timeout: {erro_timeout}")

# Verificando se há algum erro para plotar o gráfico
if erro_404 == 0 and erro_403 == 0 and erro_429 == 0 and erro_timeout == 0:
    print("Não houve erros significativos para serem plotados no gráfico.")
else:
    # Gráfico de distribuição de tipos de erros
    labels = ['Erro 404', 'Erro 403', 'Erro 429', 'Erro Timeout']
    sizes = [erro_404, erro_403, erro_429, erro_timeout]
    colors = ['#FF9999', '#66B2FF', '#99FF99', '#FFCC99']

    # Filtrando para apenas incluir valores maiores que zero
    filtered_labels = [label for label, size in zip(labels, sizes) if size > 0]
    filtered_sizes = [size for size in sizes if size > 0]
    filtered_colors = [color for color, size in zip(colors, sizes) if size > 0]

    plt.figure(figsize=(8, 8))
    plt.pie(filtered_sizes, labels=filtered_labels, colors=filtered_colors, autopct='%1.1f%%', startangle=140)
    plt.title("Distribuição dos Tipos de Erros")
    plt.axis('equal')
    plt.show()

# Análise de palavras-chave
palavras_validas = [resultado["arquivo"].split(" - ")[1].replace(".docx", "").split()
                    for resultado in resultados_validados if resultado["valido"] == True]
palavras_invalidas = [resultado["arquivo"].split(" - ")[1].replace(".docx", "").split()
                      for resultado in resultados_validados if resultado["valido"] == False]

palavras_validas_flat = [palavra for sublist in palavras_validas for palavra in sublist]
palavras_invalidas_flat = [palavra for sublist in palavras_invalidas for palavra in sublist]

# Verificando se há palavras-chave para gerar nuvem de palavras
if palavras_validas_flat:
    # Word Cloud para palavras válidas
    wc_validas = WordCloud(width=800, height=400, max_words=100, background_color='white').generate(
        " ".join(palavras_validas_flat))
    plt.figure(figsize=(10, 5))
    plt.imshow(wc_validas, interpolation='bilinear')
    plt.axis('off')
    plt.title("Palavras-Chave mais Frequentes em Registros Válidos")
    plt.show()
else:
    print("Não há palavras-chave válidas suficientes para gerar uma nuvem de palavras.")

if palavras_invalidas_flat:
    # Word Cloud para palavras inválidas
    wc_invalidas = WordCloud(width=800, height=400, max_words=100, background_color='white').generate(
        " ".join(palavras_invalidas_flat))
    plt.figure(figsize=(10, 5))
    plt.imshow(wc_invalidas, interpolation='bilinear')
    plt.axis('off')
    plt.title("Palavras-Chave mais Frequentes em Registros Inválidos")
    plt.show()
else:
    print("Não há palavras-chave inválidas suficientes para gerar uma nuvem de palavras.")

# import matplotlib.pyplot as plt
# import pickle
#
# # Carregar os resultados validados do arquivo
# with open('resultados_validados.pkl', 'rb') as f:
#     resultados_validados = pickle.load(f)
#
# # Função para gerar gráficos de validação
# def gerar_relatorio(resultados_validados):
#     # Contagem de resultados válidos e inválidos
#     validos = sum(1 for resultado in resultados_validados if resultado["valido"] == True)
#     invalidos = len(resultados_validados) - validos
#
#     # Criando um gráfico de pizza
#     labels = 'Válidos', 'Inválidos'
#     sizes = [validos, invalidos]
#     colors = ['#4CAF50', '#F44336']
#     explode = (0.1, 0)  # Destaque para válidos
#
#     plt.pie(sizes, explode=explode, labels=labels, colors=colors,
#             autopct='%1.1f%%', shadow=True, startangle=140)
#     plt.axis('equal')  # Assegura que o gráfico seja um círculo
#     plt.title("Distribuição de Resultados Válidos vs Inválidos")
#     plt.show()
#
# # Gerando o relatório
# gerar_relatorio(resultados_validados)
