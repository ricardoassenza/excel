import pandas as pd
from openpyxl import load_workbook

from umplanilha import filtra
from doisfinal import transferencia
from tresreserva import filtro_filtrada
from quatrodinamica import atualiza_dinamicas
from decrescente import ordem_decrescente


razao =  r"C:\Users\52414463899\Documents\Copia\razao.xlsx"
controle = r"C:\Users\52414463899\Documents\Copia\controle.xlsx"

filtra(razao)
transferencia(razao,controle)
filtro_filtrada(controle)
atualiza_dinamicas(controle)
ordem_decrescente(controle)

print('completed')