import pandas as pd

leitor = pd.read_excel('Pasta1.xlsx')

# 2. Preencher onde Confirmado é 'Sim'
leitor.loc[leitor['Já Chegou'] == 'Sim', 'Confirmação'] = 'Confirmado com sucesso'

# 3. Preencher onde Confirmado é 'Não'
leitor.loc[leitor['Já Chegou'] == 'Não', 'Confirmação'] = 'Contato pendente'

# 4. Salvar de volta no Excel (pode ser novo arquivo para segurança)
with pd.ExcelWriter('Pasta1.xlsx', engine='openpyxl', mode='w') as writer:
    leitor.to_excel(writer, index=False)

print('atualizada')    