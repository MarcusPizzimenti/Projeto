#!/usr/bin/env python3

import pandas as pd

# Carregar a planilha de análise
file_path_analise = r'/home/marcus_pizzi/Documents/ANALISE.xlsx'
df_cleaned = pd.read_excel(file_path_analise, sheet_name='Análise', header=4)

# Renomear colunas para facilitar o acesso
df_cleaned.columns = ['NaN' ,'Ponto', 'Área (ha)', 'pH(H2O)', 'P (mg/dm³)', 'K (mg/dm³)', 
                      'Mg (Cmolc/dm³)', 'Ca (Cmolc/dm³)', 'H+Al (Cmolc/dm³)', 'SB (Cmolc/dm³)', 
                      'CTC (Cmolc/dm³)', 'V%', 'm%']

# Dropar a coluna NaN desnecessária
df_cleaned = df_cleaned.drop(columns=['NaN'])

# Solicitar a cultura que será cultivada, repetindo a solicitação até obter uma entrada válida
valores_v2 = {'milho': 60, 'soja': 70, 'cana': 50}
cultura = ""
V2 = None

while V2 is None:
    cultura = input("Informe a cultura que será cultivada (milho, soja, cana): ").strip().lower()
    V2 = valores_v2.get(cultura)
    if V2 is None:
        print("Cultura desconhecida. Por favor, escolha entre milho, soja ou cana.\n")


# Solicitar o valor de V2 e o preço do calcário
PRNT = float(input("Insira o valor do PRNT do calcário: "))
preco_calcario = float(input("Insira o preço do calcário (R$/tonelada): "))

# Função para calcular a necessidade de calagem
def calcular_calagem(V1, CTC, V2=V2, PRNT=PRNT):
    return ((V2 - float(V1)) / (PRNT * float(CTC)))

# Criar uma lista para armazenar os resultados
resultados = []

# Variáveis para totalizar a quantidade de calcário e a área total
quantidade_total_calcario = 0
area_total = 0

# Iterar pelas linhas da planilha e calcular a NC para cada ponto
for index, row in df_cleaned.iterrows():
    try:
        if not isinstance(row['Ponto'], str) or not row['Ponto'].strip():
            continue
        
        V1 = float(row['V%'])  # Convertendo V1 para número
        CTC = float(row['CTC (Cmolc/dm³)'])  # Convertendo CTC para número
        area = float(row['Área (ha)'])  # Pegando a área do talhão

        # Calcular a necessidade de calagem (NC) em toneladas por hectare
        NC = calcular_calagem(V1, CTC)
        
        # Calcular a quantidade total de calcário para o talhão
        quantidade_calcario = NC * area
        
        # Adicionar os resultados à lista
        resultados.append([row['Ponto'], area, NC, quantidade_calcario])
        
        # Somar a quantidade total de calcário e a área total
        quantidade_total_calcario += quantidade_calcario
        area_total += area
        
        # Exibir os resultados no terminal
        print(f"Ponto: {row['Ponto']}, Área: {area:.2f} ha, Ton/ha: {NC:.2f}, Ton total: {quantidade_calcario:.2f}")
    
    except ValueError:
        continue

# Criar um DataFrame com os resultados
df_resultados = pd.DataFrame(resultados, columns=['Ponto', 'Área (ha)', 'Ton/ha', 'Ton total'])

# Adicionar a linha de total
df_total = pd.DataFrame({
    'Ponto': ['Total'],
    'Área (ha)': [area_total],
    'Ton/ha': [''],
    'Ton total': [quantidade_total_calcario]
})

# Concatenar os resultados com a linha de total
df_final = pd.concat([df_resultados, df_total], ignore_index=True)

# Carregar o arquivo de recomendações para manter os cabeçalhos
file_path_recomendacao = r'/home/marcus_pizzi/Documents/RECOMENDACAO.xlsx'
df_existing = pd.read_excel(file_path_recomendacao, sheet_name='Recomendacao', header=None)

# Adicionar os novos dados na planilha existente a partir da linha 8
with pd.ExcelWriter(file_path_recomendacao, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_final.to_excel(writer, sheet_name='Recomendacao', startrow=7, startcol=1, index=False, header=False)

# Exibir a quantidade total de calcário e o custo total
custo_total = quantidade_total_calcario * preco_calcario
print(f"\nQuantidade total de calcário necessária: {quantidade_total_calcario:.2f} toneladas")
print(f"Custo total: R$ {custo_total:.2f}")
print(f"\nResultados salvos com sucesso na planilha RECOMENDACAO.xlsx!")

