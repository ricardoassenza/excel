import pandas as pd

# Nome do arquivo e da aba de origem
arquivo = r"C:\Users\52414463899\Documents\Copia\controle.xlsx"
aba_origem = 'mar25'  # <<< coloque o nome exato da aba

# 1. Ler dados da aba específica
df = pd.read_excel(arquivo, sheet_name=aba_origem)

# 2. Filtrar apenas "Recorrente"
df_recorrente = df[df['Frequência'] == 'Recorrente']

# 3. Ordenar em ordem decrescente por "Acum. 2025"
df_ordenado = df_recorrente.sort_values(by='Acum. 2025', ascending=True)

# 4. Pegar os 24 primeiros
top24 = df_ordenado.head(24)

# 5. Agrupar o restante como "Outros"
restante = df_ordenado.iloc[24:]
soma_outros = restante['Acum. 2025'].sum()

linha_outros = pd.DataFrame([{
    'Fornecedor': 'Outros',
    'Acum. 2025': soma_outros,
    'Frequência': '',
    'Finalidade': ''
}])

# 6. Juntar os dados
df_final = pd.concat([top24, linha_outros], ignore_index=True)

# 7. Calcular porcentagem
total = df_final['Acum. 2025'].sum()
df_final['Porcentagem'] = (df_final['Acum. 2025'] / total * 100).round(2).astype(str) + '%'

# 8. Salvar em nova aba
with pd.ExcelWriter(arquivo, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df_final.to_excel(writer, sheet_name='Top24_Recorrentes', index=False)


print('ready')