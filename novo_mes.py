import pandas as pd

# Caminho do arquivo
arquivo = r"C:\Users\52414463899\Documents\Copia\controle.xlsx"

# Nomes das abas
mes_anterior = "fev25"
mes_atual = "Ordenado"

# Lê as duas abas do Excel
df_antigo = pd.read_excel(arquivo, sheet_name=mes_anterior)
df_novo = pd.read_excel(arquivo, sheet_name=mes_atual)

#  Remove espaços nos nomes das colunas
df_antigo.columns = df_antigo.columns.str.strip()
df_novo.columns = df_novo.columns.str.strip()

#  Mostra as colunas para depuração
print("Aba antiga:", df_antigo.columns.tolist())
print("Aba nova:", df_novo.columns.tolist())

#  Confere se 'Fornecedor' existe nas duas tabelas
if "Fornecedor" not in df_antigo.columns or "Fornecedor" not in df_novo.columns:
    raise ValueError("A coluna 'Fornecedor' deve existir nas duas abas.")

if "Frequência" not in df_antigo.columns:
    raise ValueError("A aba do mês anterior deve ter a coluna 'Frequência'.")

#  Faz o merge para trazer a frequência do mês anterior
df_merge = df_novo.merge(
    df_antigo[["Fornecedor", "Frequência"]],
    on="Fornecedor",
    how="left",
    suffixes=("", "_anterior")
)

#  Preenche a coluna Frequência (na nova planilha) com a anterior, se existir
df_merge["Frequência"] = df_merge["Frequência_anterior"].combine_first(df_merge["Frequência"])
df_merge.drop(columns=["Frequência_anterior"], inplace=True)

#  Se ainda houver nulos na coluna Frequência, marca como "A classificar"
df_merge["Frequência"] = df_merge["Frequência"].fillna("A classificar")

#  Salva de volta na aba do mês atual
with pd.ExcelWriter(arquivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_merge.to_excel(writer, sheet_name=mes_atual, index=False)

print(" Processamento completo.")
