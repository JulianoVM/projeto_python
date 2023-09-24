import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import os

# Leitura do arquivo CSV
df = pd.read_csv('atributos.csv', header=None)
df.columns = ["classes:", "cap-shape:", "cap-surface:", "cap-color:", "bruises:", "odor:", "gill-attachment:",
              "gill-spacing:", "gill-size:", "gill-color:", "stalk-shape:", "stalk-root:", "stalk-surface-above-ring:",
              "stalk-surface-below-ring:", "stalk-color-above-ring:", "stalk-color-below-ring:", "veil-type:",
              "veil-color:", "ring-number:", "ring-type:", "spore-print-color:", "population:", "habitat:"]

# Separação de comestíveis e não comestíveis
comestiveis = "e"
pdcome = df[df.iloc[:, 0] == comestiveis]

n_comestiveis = "p"
npdcome = df[df.iloc[:, 0] == n_comestiveis]

# Identificar o atributo de maior número em cada coluna para comestíveis e não comestíveis
atributo_maior_comes = {}
atributo_maior_ncomes = {}

for coluna in pdcome.columns[1:]:  # Começando da segunda coluna, ignorando 'classes'
    atributo_mais_probavel = pdcome[coluna].value_counts().idxmax()
    atributo_maior_comes[coluna] = atributo_mais_probavel

for coluna in npdcome.columns[1:]:  # Começando da segunda coluna, ignorando 'classes'
    atributo_mais_probavel = npdcome[coluna].value_counts().idxmax()
    atributo_maior_ncomes[coluna] = atributo_mais_probavel

# Função para criar gráfico e salvar em Excel
def criar_grafico_e_excel(df, atributos_mais_probaveis, nome_arquivo_prefixo):
    plt.figure(figsize=(10, 6))
    for coluna, atributo in atributos_mais_probaveis.items():
        quantidade = df[df[coluna] == atributo][coluna].count()
        plt.bar(coluna, quantidade, label=atributo)

    plt.xlabel('Atributos')
    plt.ylabel('Quantidade')
    plt.title(f'Maiores Probabilidades de Serem {nome_arquivo_prefixo.capitalize()}')
    plt.xticks(rotation=45)
    plt.legend(loc='best')

    # Salva o gráfico como uma imagem
    nome_arquivo_grafico = f"{nome_arquivo_prefixo}_grafico.png"
    plt.savefig(nome_arquivo_grafico, bbox_inches='tight')

    # Salva o gráfico em uma planilha Excel
    nome_arquivo_excel = f"{nome_arquivo_prefixo}_com_grafico.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active

    img = Image(nome_arquivo_grafico)
    img.anchor = "E2"
    ws.add_image(img)

    wb.save(nome_arquivo_excel)
    wb.close()

    # Exclui o arquivo temporário do gráfico
    os.remove(nome_arquivo_grafico)

    print(f"Gráfico e dados para {nome_arquivo_prefixo} salvos em arquivos Excel: {nome_arquivo_excel}")

# Criar gráfico e salvar para comestíveis
criar_grafico_e_excel(pdcome, atributo_maior_comes, "comestiveis")

# Criar gráfico e salvar para não comestíveis
criar_grafico_e_excel(npdcome, atributo_maior_ncomes, "nao_comestiveis")