import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os

# Abre janela para selecionar o arquivo
Tk().withdraw()
arquivo = askopenfilename(
    initialdir=r"C:\Users\felip\Desktop\Python FJ",
    filetypes=[("Excel files", "*.xlsx")]
)

# Lê a planilha com cabeçalho na linha 4 (índice 3)
df_original = pd.read_excel(arquivo, header=3)

# Verifica se as colunas esperadas existem
colunas_esperadas = ['Nome', 'Data', 'Motivo']
if not all(col in df_original.columns for col in colunas_esperadas):
    raise ValueError("As colunas 'Nome', 'Data' e 'Motivo' devem estar presentes.")

# Remove espaços e trata "Motivo" como string
df_original['Motivo'] = df_original['Motivo'].astype(str).str.strip()

# Filtra linhas com "Motivo" efetivamente vazio ('' ou 'nan' stringificados)
df_filtrado = df_original[df_original['Motivo'].isin(['', 'nan'])]

# Gera mensagens por nome
mensagens = []
for nome, grupo in df_filtrado.groupby('Nome'):
    datas = grupo['Data'].astype(str).tolist()
    if datas:
        texto_datas = "\n".join(datas)
        mensagem = (
            f"Olá {nome}, poderia justificar as suas horas extras dos seguintes dias:\n"
            f"{texto_datas}\n\n"
            "Atenciosamente,\nGestão de Ponto"
        )
        mensagens.append({
            'Nome': nome,
            'Datas pendentes': texto_datas,
            'Mensagem': mensagem
        })

# Converte para DataFrame
df_mensagens = pd.DataFrame(mensagens)

# Caminho de saída
saida = os.path.join(os.path.dirname(arquivo), "mensagens_com_planilha.xlsx")

# Escreve a planilha com duas abas
with pd.ExcelWriter(saida, engine='openpyxl') as writer:
    df_original.to_excel(writer, index=False, sheet_name='Planilha Original')
    df_mensagens.to_excel(writer, index=False, sheet_name='Mensagens')

print(f"✅ Arquivo gerado com sucesso: {saida}")
