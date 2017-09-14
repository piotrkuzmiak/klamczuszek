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
import xlsxwriter
import pdb

class Excel_report():
    
    """
    Klasa dla przygotowania raportu w Excel w oparciu o Dataframe
    """
        
    def __init__(self, dataframe, groupby):
        """
        Konstruktor do utworzenia obiektu dla wyeksportowania raportu z podliczeniem.
        
        Parameters
        ----------
        dataframe: Dataframe
            zrodlo dla raportu w postaci Dataframe
        groupby: list[str]
            lista kolumn po ktorej przeprowadzic grupowanie na pandas.Dataframe()
            
        Returns
        -------
        Object
                   
        Notes:
        -----
        """
        self.dataframe = dataframe
        self.no_columns = len(dataframe.columns)
        self.groups = groupby
        
        
        
    def unload(self, path, sheet_name):
        """
        Funkcja do weksportowania pandas.Dataframe() do Excel'a.
        Wyesportowany arkusz zawiera grupowanie.
        
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
        Plik w formacie xlsx
        """
        
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            rowindex = 0
            for x in range(1,len(self.groups)+1):
                for name, group in self.dataframe.groupby(by=self.groups[:x]):
                    no_rows = group.shape[0]
                    group.to_excel(writer, sheet_name=sheet_name, startrow=rowindex, header=False)
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet_name]
                    for index in group.index.tolist():
#                        pdb.set_trace()
                        worksheet.set_row(index, None, None, {'level': x})
                    rowindex = no_rows+1
                    
#                worksheet.set_column('A:C', 10)
#            worksheet.set_column('A:C", 10)