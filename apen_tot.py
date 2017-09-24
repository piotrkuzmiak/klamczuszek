#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 23 14:58:05 2017

@author: piotr
"""
import pandas as pd
import pdb

df = pd.DataFrame([['zachod','129175','sprzedaz','Poznan',1, 1],
                   ['zachod','129175','sprzedaz','Poznan',1, 1],\
                   ['wschod','118158','sprzedaz','Lublin', 10, 20],\
                   ['wschod','114678','konto','Warszawa', 10, 10],\
                   ['wschod','114678','sprzedaz','Warszawa', 10, 6],\
                   ['zachod','129175','sprzedaz','Poznan',1,1],\
                   ['zachod','129175','konto','Poznan',90,4],\
                   ['zachod','129175','sprzedaz','Poznan',5,2],\
                   ['zachod','129175','sprzedaz','Poznan',4,5],
['poludnie','130115','konto','Krakow',4,4]], columns=['makroregion','skp','kampania','oddzial' ,'khd_info', 'billings'])
#%%
def append_tot(df):
    """
    
    Author:
    ------
    piRSquared
    (Stackoverflow Aug 15 '16 at 23:42)
    
    """
#    pdb.set_trace()
    if hasattr(df, 'name') and df.name is not None:
        xs = df.xs(df.name)
    else:
        xs = df
    gb = xs.groupby(level=0)
    n = xs.index.nlevels
    name = tuple('Total' if i == 0 else '' for i in range(n))
    tot = gb.sum().sum().rename(name).to_frame().T
    if n > 1:
        sm = gb.apply(append_tot)
    else:
        sm = gb.sum()
    return pd.concat([sm, tot])

fields = ['makroregion','oddzial','kampania','skp']
df=append_tot(df.set_index(fields)).reset_index()
#df['procent']=df['billings']/df['khd_info']
df