import pandas as pd
import matplotlib.pyplot as plt

# Carregar dados do pickle
df = pd.read_pickle("dados_clipping.pkl")

# Verificar as colunas disponíveis
print("Colunas disponíveis no DataFrame:", df.columns)

# Escolher uma coluna para criar o gráfico de pizza
coluna_para_grafico = 'ASSUNTO'

# Verificar se a coluna existe no DataFrame
if coluna_para_grafico not in df.columns:
    print(f"Coluna '{coluna_para_grafico}' não encontrada no DataFrame.")
else:
    # Agrupar os dados e contar as ocorrências
    distribuicao_coluna = df[coluna_para_grafico].value_counts()

    # Ajustar para agrupar categorias menos representativas
    limite = 2  # Definir o limite para agrupar categorias menores
    outros = distribuicao_coluna[distribuicao_coluna <= limite].sum()
    distribuicao_coluna = distribuicao_coluna[distribuicao_coluna > limite]
    distribuicao_coluna['Outros'] = outros

    # Criar gráfico de pizza
    plt.figure(figsize=(12, 8))  # Aumentar o tamanho do gráfico para 12x8 polegadas

    # Definir cores e explosão de fatias
    colors = plt.cm.tab20.colors
    explode = [0.1 if i < 5 else 0 for i in range(len(distribuicao_coluna))]

    # Plotar gráfico de pizza com melhorias
    patches, texts, autotexts = plt.pie(
        distribuicao_coluna,
        labels=distribuicao_coluna.index,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        explode=explode,
        pctdistance=0.85  # Distância dos números para o centro
    )

    # Ajustar tamanho e estilo dos textos
    for text in texts:
        text.set_fontsize(10)
        text.set_bbox(dict(facecolor='white', edgecolor='none', alpha=0.6))  # Fundo branco para rótulos
    for autotext in autotexts:
        autotext.set_fontsize(12)
        autotext.set_color('black')  # Números em preto
        autotext.set_weight('bold')  # Negrito nos números

    # Adicionar título
    plt.title(f'Distribuição de Publicações por {coluna_para_grafico}', fontsize=16)

    # Desenhar um círculo no centro para "cortar" o gráfico de pizza e torná-lo um "donut"
    centre_circle = plt.Circle((0, 0), 0.70, fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)

    # Garantir que o gráfico seja exibido como um círculo
    plt.gca().set_aspect('equal')

    # Adicionar legenda
    plt.legend(patches, distribuicao_coluna.index, loc='best', bbox_to_anchor=(1, 0.5), fontsize=10)

    # Remover rótulo do eixo Y e ajustar layout
    plt.ylabel('')
    plt.tight_layout()
    plt.show()
