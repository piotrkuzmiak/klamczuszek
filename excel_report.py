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
        
        
        
    def unload(self, path, sheet_name='Arkusz1'):
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
            workbook  = writer.book
            format_percentage = workbook.add_format({'num_format': '0%'})
            rowindex = 0
            for name_mc, makro in self.dataframe.groupby(by=self.groups[0]):
                print((name_mc))
                df_sk_makro = makro[['khd_info','billings']].sum()
                rowindex=rowindex+df_sk_makro.shape[0]
                for name_reg, region in makro.groupby(by=self.groups[1]):
                    print(name_reg)
                    df_sk_reg = region[['khd_info','billings']].sum()
                    rowindex=rowindex+df_sk_reg.shape[0]
                    for name_od, od in region.groupby(by=self.groups[2]):
                        print(name_od)
                        df_sk_od = od[['khd_info','billings']].sum()
                        rowindex=rowindex+df_sk_od.shape[0]
                        for name_skp, skp in od.groupby(by=self.groups[3]):
                            print(name_skp)
                            pdb.set_trace()
                            df_sk_sum = skp[['khd_info','billings']].sum(numeric_only=True)
                            rowindex=rowindex+df_sk_sum.shape[0]
                            df_sk_sum.to_excel(writer, sheet_name=sheet_name, header=False, startrow=rowindex)
                            worksheet = writer.sheets[sheet_name]
                            print('podsumowanie dla: ',name_skp,df_sk_sum)
                        print('podsumowanie dla: ',name_od,df_sk_od)
                    print('podsumowanie dla: ', name_reg, df_sk_reg)
                print('podsumowanie dla: ', name_mc,df_sk_makro )
                df_sk_makro.to_excel(writer, sheet_name=sheet_name, header=False)
#                worksheet = writer.sheets[sheet_name]
                
                
    def unload_pivot(self, path, sheet_name='Arkusz1'):
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            workbook  = writer.book
            format_percentage = workbook.add_format({'num_format': '0%'})
            pv = self.dataframe.pivot_table(values=['khd_info','billings','% oznaczonych kontaktów'],\
                                       index=['makroregion', 'region', 'oddzial', 'SKP', 'NAZWA_KAMPANII','offer_type_cd'],\
                                       aggfunc=sum).reset_index()
            pv.to_excel(writer,sheet_name=sheet_name, header=True, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('G:G',30,format_percentage)
            worksheet.set_column('A:C',35,None)
            worksheet.set_column('E:E',50,None)
            worksheet.autofilter(0,0,pv.shape[0],pv.shape[1])