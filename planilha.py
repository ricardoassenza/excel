import pandas as pd

caminho_arquivo = r"C:\Users\52414463899\Documents\Copia\03-Razão_Energia-Mar25_copia.xlsx"

leitor = pd.read_excel(caminho_arquivo)

itens_coluna_V =['Benefícios', 'Custo de Energia Submercado', 'Custo de Energia Padrão',  'Custo de Energia Mercado Curto Prazo - MCP','Eficiência tributária de compra de energia', 'Encargos', 'Encargos de Conexão','Entidade de Classe','Recomposição de Danos Patrimoniais','Salários','Seguros','Taxa Fiscaliz. TFSEE/ONS/CCEE','Taxas e Impostos','TUST']

itens_coluna_E = ['GD - PECUARIA SERRAMAR', 'GD - BRASILIA', 'GD - SANTA RITA', 'SERVENG ENERGIAS IMOBILIARIA']

filtro_final = leitor[
    (leitor['DescDRE'] == 'Custos e Despesas Operacionais') &
    (leitor['AgrupadorTipoRateio'] == 'Operacional') &
    (leitor['Periodo'] == 2503) &
    (~leitor['DescRubrica'].isin(itens_coluna_V)) &
    (leitor['NomeCliFor'].notna()) &
    (leitor['DescricaoFilial'].notna()) &
    (~leitor['DescricaoFilial'].isin(itens_coluna_E))
    ]

with pd.ExcelWriter(caminho_arquivo, engine='openpyxl', mode='a') as writer:
    filtro_final.to_excel(writer, sheet_name='base_filtrada', index=False)

print('feito')


###############################################################################################################





