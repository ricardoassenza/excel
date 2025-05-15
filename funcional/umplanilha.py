import pandas as pd


def filtra(caminho):

    leitor = pd.read_excel(caminho)

    itens_coluna_V =['Benefícios', 'Custo de Energia Submercado', 'Custo de Energia Padrão',  'Custo de Energia Mercado Curto Prazo - MCP','Eficiência tributária de compra de energia', 'Encargos', 'Encargos de Conexão','Entidade de Classe','Recomposição de Danos Patrimoniais','Salários','Seguros','Taxa Fiscaliz. TFSEE/ONS/CCEE','Taxas e Impostos','TUST']

    itens_coluna_E = ['GD - PECUARIA SERRAMAR', 'GD - BRASILIA', 'GD - SANTA RITA', 'SERVENG ENERGIAS IMOBILIARIA']

    filtro_final = leitor[
        (leitor['Periodo'] == 2503) &
        (leitor['DescDRE'] == 'Custos e Despesas Operacionais') &
        (leitor['AgrupadorTipoRateio'] == 'Operacional') &
        (~leitor['DescRubrica'].isin(itens_coluna_V)) &
        (leitor['NomeCliFor'].notna()) &
        (leitor['DescricaoFilial'].notna()) &
        (~leitor['DescricaoFilial'].isin(itens_coluna_E))
        ]

    with pd.ExcelWriter(caminho, engine='openpyxl', mode='a') as writer:
        filtro_final.to_excel(writer, sheet_name='base_filtrada', index=False)

    print('filtros feitos')


###############################################################################################################





