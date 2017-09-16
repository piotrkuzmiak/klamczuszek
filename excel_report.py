# -*- coding: utf-8 -*-
"""
Created on Thu Aug 10 14:45:13 2017

@author: 129175
"""

#TODO: najpierw przygotowanie raportu:
#- polaczenie z khd i pobranie danych za dany miesiac (z zapytaniem za jaki miesiac),
#- przygotowanei raportu w excel,
#- zapisanie df_all_connections do playpaen na khd
#- zrzucic raport na zadana lokalizacje i powiadomienie ze raport jest dostepny (control-m)



#%%
import pandas as pd
import pdb

class Excel_report():
    
    """
    Klasa dla przygotowania raportu w Excel w oparciu o Dataframe
    """
        
    def __init__(self, dataframe, groupby):
        """
        Konstruktor do utworzenia obiektu dla wyeksportowania raportu z
        podliczeniem.
        
        Parameters
        ----------
        dataframe: Dataframe
            zrodlo dla raportu w postaci Dataframe
        groupby: list[str]
            lista kolumn po ktorej przeprowadzic grupowanie na 
            pandas.Dataframe()
            
        Returns
        -------
        Object
                   
        Notes:
        -----
        """
        self.dataframe = dataframe
        self.no_columns = len(dataframe.columns)
        self.groups = groupby
        
        
        
    def unload(self, path, sheet_name='Arkusz1'):
        """
        Funkcja do weksportowania pandas.Dataframe() do Excel'a z podsumowaniem
        w grupach
                
        Parameters
        ----------
        path: str
            sciezka do pliku wynikowego z nazwa samego pliku Excel w sciezce.
        sheet_name: str
            Nazwa arkusza w pliku Excel
            
        Returns
        -------
        Excel file:Excel object
                   
        Notes:
        -----
        """
        
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            workbook  = writer.book
            format_percentage = workbook.add_format({'num_format': '0%'})
            df = self._append_tot(self.dataframe.set_index(self.groups)).reset_index(col_fill='index')
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.autofilter(0,0,df.shape[0],df.shape[1]+1-len(self.groups))
                
    def unload_pivot(self, path, sheet_name='Arkusz1'):
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            workbook  = writer.book
            format_percentage = workbook.add_format({'num_format': '0%'})
            pv = self.dataframe.pivot_table(values=['khd_info','billings',\
                                                    '% oznaczonych kontaktów'],\
                                       index=['makroregion', 'region',\
                                              'oddzial', 'SKP', 'NAZWA_KAMPANII',\
                                              'offer_type_cd'],\
                                       aggfunc=sum).reset_index()
            pv.to_excel(writer,sheet_name=sheet_name, header=True, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('G:G',30,format_percentage)
            worksheet.set_column('A:C',35,None)
            worksheet.set_column('E:E',50,None)
            worksheet.autofilter(0,0,pv.shape[0],pv.shape[1])
            
    def _append_tot(self,df_new_index):
        """
        Funkcja do generowania pandas.Dataframe'a z podsumowaniami w grupach
        
        Parameters
        ----------
        df_new_index: Dataframe
            zrodlo dla raportu w postaci pandas.Dataframe() z nadaniem nowego
            indeksu
            
        Returns
        -------
        pandas.Dataframe()
        
        
        Example:
        --------
        fields = ['makroregion','oddzial','skp','kampania']
        append_tot(df.set_index(fields)).reset_index()
                
        
        Author:
        -------
        piRSquared
        (Stackoverflow Aug 15 '16 at 23:42)
        
        """
        
        if hasattr(df_new_index, 'name') and df_new_index.name is not None:
            xs = df_new_index.xs(df_new_index.name)
        else:
            xs = df_new_index
        gb = xs.groupby(level=0)
        n = xs.index.nlevels
        name = tuple('Suma' if i == 0 else '' for i in range(n))
        tot = gb.sum().sum().rename(name).to_frame().T
        if n > 1:
            sm = gb.apply(self._append_tot)
        else:
            sm = gb.sum()
        return pd.concat([sm, tot])
