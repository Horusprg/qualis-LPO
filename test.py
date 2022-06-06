import xlrd
import pandas as pd
from collections import Counter
import numpy as np
from xlutils.copy import copy

rb = xlrd.open_workbook("Resultado.xls")        #Ler arquivo para fazer cópia
wb = copy(rb)
w_sheet = wb.get_sheet(0)
df = pd.read_excel("Resultado.xls")
df2 = pd.read_csv('qualis.csv')
R = list(Counter(df['Titulo Periodico ou Revista']))[1:]
w_sheet.write(0, 11, "Estrato Antigo")

dic = {}

for revista in R:
    idxrevista = df2[df2['Título'] == revista.upper()].index
    if len(idxrevista) > 0:
        estratoantigo = df2.loc[idxrevista]['Estrato'][idxrevista[0]]
    else:
        estratoantigo = ''
        
    dic[revista] = estratoantigo

for revista in dic:
    idxs = df[df["Titulo Periodico ou Revista"] == revista].index
    w_sheet.write(idxs.values[0], 11, dic[revista])

wb.save("Resultado.xls")