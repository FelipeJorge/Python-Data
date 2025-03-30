import pandas as pd

# Carregar o arquivo Excel, considerando que os cabeçalhos estão na linha 4 (índice 3)
df = pd.read_excel("C:\\Users\\felip\\Desktop\\testeponto.xlsx", header=3)

# Exibir os nomes das colunas para verificar o índice correto
print("Colunas originais:", df.columns)

df["Colaborador"] = None  

# Variável para armazenar o nome atual do colaborador
colaborador_atual = None

# Percorrer as linhas e identificar quando o nome do colaborador muda
for index, row in df.iterrows():
    if "Colaborador" in str(row[0]):  # Se a palavra "Colaborador" estiver na primeira coluna
        colaborador_atual = row[1]  # Captura o nome do colaborador da coluna B
        df.at[index, "Colaborador"] = colaborador_atual  # Define o nome na mesma linha
    
    else:
        df.at[index, "Colaborador"] = colaborador_atual  # Preenche as linhas subsequentes com o nome do colaborador atual

df = df[~df.iloc[:, 0].astype(str).str.contains("Colaborador", na=False)]

colunas_ordenadas = ["Colaborador"] + [col for col in df.columns if col != "Colaborador"]
df = df[colunas_ordenadas]

# Definir as colunas de horário
colunas_horario = ['1ª Entrada', '1ª Saída', '2ª Entrada', '2ª Saída']

# Criar uma nova coluna com a contagem de preenchimentos em cada linha
df["Preenchidos"] = df[colunas_horario].notna().sum(axis=1)

# Verificar se a coluna "Motivo/Observação" está vazia
filtro_motivo = df["Motivo/Observação"].isna() | (df["Motivo/Observação"].astype(str).str.strip() == "")

# Filtrar as linhas que não são sábado e têm menos de 4 preenchimentos **e** "Motivo/Observação" vazio
filtro_geral = (df["Preenchidos"] < 4) & (~df["Data"].astype(str).str.startswith("Sáb")) & filtro_motivo

# Filtrar os sábados que têm menos de 2 preenchimentos **e** "Motivo/Observação" vazio
filtro_sabado = (df["Preenchidos"] < 2) & (df["Data"].astype(str).str.startswith("Sáb")) & filtro_motivo

# Aplicar os filtros combinados
df_filtrado = df[filtro_geral | filtro_sabado]

# Remover a coluna auxiliar "Preenchidos"
df_filtrado = df_filtrado.drop(columns=["Preenchidos"])

# Exibir o resultado
print(df_filtrado)

# Salvar o resultado em um novo arquivo Excel (se necessário)
df_filtrado.to_excel("C:\\Users\\felip\\Desktop\\ponto_faltantes.xlsx", index=False)





