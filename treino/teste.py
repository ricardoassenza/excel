import pandas as pd

arquivo = r'treino\vendas_com_status.xlsx'

fonte = 'dados'
destino = 'Extra'

aba_fonte = pd.read_excel(arquivo, sheet_name=fonte)
aba_destino = pd.read_excel(arquivo, sheet_name=destino)

aba_fonte.columns = aba_fonte.columns.str.strip()
aba_destino.columns = aba_destino.columns.str.strip()

print(f'titulos da aba dados: {aba_fonte.columns.to_list()}')
print(f'titulos da aba Extra: {aba_destino.columns.to_list()}')

if 'Produto' not in aba_fonte.columns or 'Produto' not in aba_destino.columns:
    raise ValueError('Titulo não encontrado')

if 'teste' not in aba_fonte.columns:
    raise ValueError('teste deve ter na aba dados')

junta_abas = aba_destino.merge(
    aba_fonte[['Produto', 'teste']],
    on='Produto',
    how='left',
    suffixes=('','_usada')
) 

junta_abas['teste'] = junta_abas['teste_usada'].combine_first(junta_abas['teste'])
junta_abas.drop(columns=['teste_usada'], inplace=True)

with pd.ExcelWriter(arquivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as wirter:
    junta_abas.to_excel(wirter, sheet_name='Extra', index=False)

print('executado')    