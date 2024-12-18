#!usr/bin/env python3

import pandas as pd

# Definir PRNT aqui, por exemplo:
PRNT = 90  # Ajuste conforme necessário

# Custos
custo_cal = 125  # Custo por tonelada de calcario calcitico, em reais (2024)
custo_ges = 140  # Custo por tonelada de gesso agricola, em reais (2024)


# Funções para calcular a necessidade de calagem e gessagem
def calcular_calagem(V1, V2, CTC, area):
    # Cálculo da necessidade de calagem
    return (((V2 - V1) * CTC) / (10 * PRNT)) * area

def calcular_gessagem(V1, V2, CTC, area):
    # Cálculo da necessidade de gessagem
    return (((V2 - V1) * CTC) / 500) * area

# Ler o arquivo Excel
df = pd.read_excel(r'/home/marcus_pizzi/Documents/Analise.xlsx', 
                   sheet_name='Análise química de solo',
                   skiprows=[1, 2])

# Garantir que as colunas relevantes sejam numéricas
df['m'] = pd.to_numeric(df['m'], errors='coerce')
df['Ca'] = pd.to_numeric(df['Ca'], errors='coerce')
df['V'] = pd.to_numeric(df['V'], errors='coerce')
df['CTC'] = pd.to_numeric(df['CTC'], errors='coerce')
df['Área talhão'] = pd.to_numeric(df['Área talhão'], errors='coerce')
df['Pontos'] = pd.to_numeric(df['Pontos'], errors='coerce')

# Lista para armazenar resultados
resultados = []

# Analisando os dados
for index, row in df.iterrows():
    area = row["Área talhão"]
    V1 = row["V"]
    V2 = 70 if row["Cultura"] == "Milho" else 60  # V2 baseado na cultura
    CTC = row["CTC"]
    m = row["m"]
    Ca = row["Ca"]
    profundidade = row["Profun"]  # Supondo que esta seja a coluna com profundidade
    Pontos = row["Pontos"]
    
    NG = 0
    NC = 0


    # Lógica com base na profundidade
    if profundidade == "0-20":
        if V1 < 50:  # Calagem se V% < 50%
            NC = calcular_calagem(V1, V2, CTC, area)
    elif profundidade == "20-40":
        if m > 20 or Ca < 5 or V1 < 35:  # Condições para gessagem
            NG = calcular_gessagem(V1, V2, CTC, area)

      # Cálculo de custos
    custo_total_calagem = NC * custo_cal
    custo_total_gessagem = NG * custo_ges


    # Adicionando resultados
    resultados.append({"Pontos": row["Pontos"], "Cultura": row["Cultura"], "Tipo": "Calagem", "Área talhão": area, "Quantidade (t)": NC, "Custo Total (R$)": custo_total_calagem})
    resultados.append({"Pontos": row["Pontos"], "Cultura": row["Cultura"], "Tipo": "Gessagem", "Área talhão": area, "Quantidade (t)": NG, "Custo Total (R$)": custo_total_gessagem})

# Criando um DataFrame com os resultados
resultados_df = pd.DataFrame(resultados, columns=["Pontos", "Cultura", "Tipo", "Área talhão", "Quantidade (t)", "Custo Total (R$)"])

# Salvando os resultados em um arquivo CSV
resultados_df.to_csv(r'/home/marcus_pizzi/Documents/resultados.csv', index=False)

# Exibindo resultados
print("Resultados salvos em resultados.csv")

