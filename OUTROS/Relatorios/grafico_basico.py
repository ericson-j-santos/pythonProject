import pandas as pd
import matplotlib.pyplot as plt

# Carregar dados do pickle
df = pd.read_pickle("dados_clipping.pkl")

# Agrupar os dados por veículo de comunicação e contar as ocorrências
top_veiculos = df['VEÍCULO'].value_counts().head(10)

# Criar o gráfico
plt.figure(figsize=(10, 6))
top_veiculos.plot(kind='bar', color='skyblue')
plt.title('Top 10 Veículos de Comunicação')
plt.xlabel('Veículo')
plt.ylabel('Número de Publicações')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()
