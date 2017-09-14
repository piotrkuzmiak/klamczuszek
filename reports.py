# -*- coding: utf-8 -*-
"""
Created on Fri Sep  8 15:44:16 2017

@author: 129175
"""
import pandas as pd
import numpy as np
from excel_report import Excel_report

<<<<<<< HEAD
with sqlite3.connect('piotr.db') as connection:
    df = pd.read_sql_query('select obreb, funkcja_dominujaca from nieruchomosci', con=connection)

def generuj(lista):
    return choice(lista)

#excel= Excel_report(dataframe=df, groupby='obreb')
#excel.unload(path='Raport.xlsx', sheet_name='arkusz1')
df['mobile'] = np.NaN
df['wired'] = np.NaN
df['SKP'] = np.nan
df.mobile = df.mobile.apply(lambda x : generuj(lista=[1,2,3, np.nan]))
df.wired = df.wired.apply(lambda x : generuj(lista=[1,2,3, np.nan]))
df.SKP = df.SKP.apply(lambda x : generuj(lista=[129175,116110,101900,120120,\
                                                130130,111111,222222]))
df_biling = pd.DataFrame(np.random.choice([1,2,3,6,7,8,9], size=(50,2)), columns=['numer_a','numer_b'])    
df_biling['SKP'] = np.random.choice([129175,116110,101900,120120,130130,111111,222222], size=(50,1))

excel = Excel_report(df[['obreb','funkcja_dominujaca','mobile','wired']], groupby=['obreb','funkcja_dominujaca'])
excel.unload(path='klamczuch.xlsx', sheet_name='Raport')
=======
df = pd.DataFrame([[129175,'sprzedaz',True, False],[118158,'sprzedaz', False, False],[129175,'sprzedaz',True,True]], columns=['skp','kampania', 'rejestr', 'biling'])
df.loc[:,['rejestr', 'biling']] = df[['rejestr', 'biling']].astype(np.int)
excel = Excel_report(dataframe=df)
excel.unload('D:/CRM/wymiana/klamczuszek.xlsx','Raport',groupby=['skp', 'kampania'])
>>>>>>> 50246dd0657a035b74e02799ad6e4b09871aa170
