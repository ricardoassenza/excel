import pandas as pd
from openpyxl import load_workbook

from umplanilha import filtra
from doisfinal import transferencia
from tresreserva import filtro_filtrada
from quatrodinamica import atualiza_dinamicas
from decrescente import ordem_decrescente
from novo_mes import frequencia, finalidade

razao =  r"C:\Users\52414463899\Documents\Copia\razao.xlsx"
controle = r"C:\Users\52414463899\Documents\Copia\controle.xlsx"

'''filtra(razao)
transferencia(razao,controle)
filtro_filtrada(controle)
atualiza_dinamicas(controle)'''
ordem_decrescente(controle, 'mar25')
frequencia(controle, 'fev25' ,'mar25')
finalidade(controle, 'fev25' ,'mar25')

print('completed')