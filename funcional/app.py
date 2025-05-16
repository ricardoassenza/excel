import pandas as pd
from openpyxl import load_workbook

from umplanilha import filtra
from doisfinal import transferencia
from tresreserva import filtro_filtrada
from quatrodinamica import atualiza_dinamicas
from decrescente import ordem_decrescente
from novo_mes import frequencia, finalidade
from ultima import recorrente, nao_recorrente
from contabil import formatacao
from aparencia import pintura


razao =  r"C:\Users\52414463899\Documents\Copia\RAZAO.xlsx"
controle = r"C:\Users\52414463899\Documents\Copia\Controle.xlsx"


#funcional:
filtra(razao)
transferencia(razao,controle)
filtro_filtrada(controle)
atualiza_dinamicas(controle)
ordem_decrescente(controle, 'abr25')
frequencia(controle, 'Fonte-Dados' ,'abr25')
finalidade(controle, 'Fonte-Dados' ,'abr25')
recorrente(controle, 'abr25', 'Recorrente')
nao_recorrente(controle, 'abr25', 'Não Recorrente')
formatacao(controle, 'abr25', 'B')
formatacao(controle, 'Recorrente', 'B')
formatacao(controle, 'Não Recorrente', 'B')

#estetico:
pintura(controle, 'abr25', 'FF9900', '000000' )
pintura(controle, 'Recorrente', '0F243E', 'FFFFFF' )
pintura(controle, 'Não Recorrente', '0F243E', 'FFFFFF' )


print('completed')