# -*- coding: utf-8 -*-
"""
Created on Fri Sep  8 15:44:16 2017

@author: 129175
"""
import pandas as pd
import numpy as np
from excel_report import Excel_report

df = pd.DataFrame([[129175,'sprzedaz',True, False],[118158,'sprzedaz', False, False],[129175,'sprzedaz',False,True]], columns=['skp','kampania', 'rejestr', 'biling'])
df.loc[:,['rejestr', 'biling']] = df[['rejestr', 'biling']].astype(np.int)
excel = Excel_report(dataframe=df)
excel.unload('D:/CRM/wymiana/klamczuszek.xlsx','Raport',groupby=['skp', 'kampania'])