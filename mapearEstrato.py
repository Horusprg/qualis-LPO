import pandas as pd
from collections import Counter
import numpy as np
import xlrd

def mapear():
    df = pd.read_excel('Resultado.xls')
    df2 = pd.read_excel('qualis.xls')
    R = list(Counter(df['Titulo Periodico ou Revista']))[1:]

    dic = {}

    for revista in R:
        idxrevista = df2[df2['TÍTULO'] == revista.upper()].index
        
        if len(idxrevista) > 0:
            estratoantigo = df2.loc[idxrevista]['ESTRATO'][idxrevista[0]]
        else:
            estratoantigo = np.nan
            
        dic[revista] = estratoantigo

    for revista in dic:
        idxs = df[df["Titulo Periodico ou Revista"] == revista].index
        df.loc[idxs, "Estrato Antigo"] = dic[revista]

    df.to_excel("Resultado.xls")
