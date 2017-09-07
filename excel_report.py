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



###############################################################################
#
# Example of how use Python and XlsxWriter to generate Excel outlines and
# grouping.
#
# Excel allows you to group rows or columns so that they can be hidden or
# displayed with a single mouse click. This feature is referred to as outlines.
#
# Outlines can reduce complex data down to a few salient sub-totals or
# summaries.
#
# Copyright 2013-2017, John McNamara, jmcnamara@cpan.org
#
#%%
import xlsxwriter

# Create a new workbook and add some worksheets
workbook = xlsxwriter.Workbook('outline.xlsx')
worksheet1 = workbook.add_worksheet('Outlined Rows')
worksheet2 = workbook.add_worksheet('Collapsed Rows')
worksheet3 = workbook.add_worksheet('Outline Columns')
worksheet4 = workbook.add_worksheet('Outline levels')

# Add a general format
bold = workbook.add_format({'bold': 1})


###############################################################################
#
# Example 1: A worksheet with outlined rows. It also includes SUBTOTAL()
# functions so that it looks like the type of automatic outlines that are
# generated when you use the Excel Data->SubTotals menu item.
#
# For outlines the important parameters are 'level' and 'hidden'. Rows with
# the same 'level' are grouped together. The group will be collapsed if
# 'hidden' is enabled. The parameters 'height' and 'cell_format' are assigned
# default values if they are None.
#
worksheet1.set_row(1, None, None, {'level': 2})
worksheet1.set_row(2, None, None, {'level': 2})
worksheet1.set_row(3, None, None, {'level': 2})
worksheet1.set_row(4, None, None, {'level': 2})
worksheet1.set_row(5, None, None, {'level': 1})

worksheet1.set_row(6, None, None, {'level': 2})
worksheet1.set_row(7, None, None, {'level': 2})
worksheet1.set_row(8, None, None, {'level': 2})
worksheet1.set_row(9, None, None, {'level': 2})
worksheet1.set_row(10, None, None, {'level': 1})

# Adjust the column width for clarity
worksheet1.set_column('A:A', 20)

# Add the data, labels and formulas
worksheet1.write('A1', 'Region', bold)
worksheet1.write('A2', 'North')
worksheet1.write('A3', 'North')
worksheet1.write('A4', 'North')
worksheet1.write('A5', 'North')
worksheet1.write('A6', 'North Total', bold)

worksheet1.write('B1', 'Sales', bold)
worksheet1.write('B2', 1000)
worksheet1.write('B3', 1200)
worksheet1.write('B4', 900)
worksheet1.write('B5', 1200)
worksheet1.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

worksheet1.write('A7', 'South')
worksheet1.write('A8', 'South')
worksheet1.write('A9', 'South')
worksheet1.write('A10', 'South')
worksheet1.write('A11', 'South Total', bold)

worksheet1.write('B7', 400)
worksheet1.write('B8', 600)
worksheet1.write('B9', 500)
worksheet1.write('B10', 600)
worksheet1.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

worksheet1.write('A12', 'Grand Total', bold)
worksheet1.write('B12', '=SUBTOTAL(9,B2:B10)', bold)


###############################################################################
#
# Example 2: A worksheet with outlined rows. This is the same as the
# previous example except that the rows are collapsed.
# Note: We need to indicate the rows that contains the collapsed symbol '+'
# with the optional parameter, 'collapsed'. The group will be then be
# collapsed if 'hidden' is True.
#
worksheet2.set_row(1, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(2, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(3, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(4, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(5, None, None, {'level': 1, 'hidden': True})

worksheet2.set_row(6, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(7, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(8, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(9, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(10, None, None, {'level': 1, 'hidden': True})
worksheet2.set_row(11, None, None, {'collapsed': True})

# Adjust the column width for clarity
worksheet2.set_column('A:A', 20)

# Add the data, labels and formulas
worksheet2.write('A1', 'Region', bold)
worksheet2.write('A2', 'North')
worksheet2.write('A3', 'North')
worksheet2.write('A4', 'North')
worksheet2.write('A5', 'North')
worksheet2.write('A6', 'North Total', bold)

worksheet2.write('B1', 'Sales', bold)
worksheet2.write('B2', 1000)
worksheet2.write('B3', 1200)
worksheet2.write('B4', 900)
worksheet2.write('B5', 1200)
worksheet2.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

worksheet2.write('A7', 'South')
worksheet2.write('A8', 'South')
worksheet2.write('A9', 'South')
worksheet2.write('A10', 'South')
worksheet2.write('A11', 'South Total', bold)

worksheet2.write('B7', 400)
worksheet2.write('B8', 600)
worksheet2.write('B9', 500)
worksheet2.write('B10', 600)
worksheet2.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

worksheet2.write('A12', 'Grand Total', bold)
worksheet2.write('B12', '=SUBTOTAL(9,B2:B10)', bold)


###############################################################################
#
# Example 3: Create a worksheet with outlined columns.
#
data = [
    ['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total'],
    ['North', 50, 20, 15, 25, 65, 80, '=SUM(B2:G2)'],
    ['South', 10, 20, 30, 50, 50, 50, '=SUM(B3:G3)'],
    ['East', 45, 75, 50, 15, 75, 100, '=SUM(B4:G4)'],
    ['West', 15, 15, 55, 35, 20, 50, '=SUM(B5:G5)']]

# Add bold format to the first row.
worksheet3.set_row(0, None, bold)

# Set column formatting and the outline level.
worksheet3.set_column('A:A', 10, bold)
worksheet3.set_column('B:G', 5, None, {'level': 1})
worksheet3.set_column('H:H', 10)

# Write the data and a formula
for row, data_row in enumerate(data):
    worksheet3.write_row(row, 0, data_row)

worksheet3.write('H6', '=SUM(H2:H5)', bold)


###############################################################################
#
# Example 4: Show all possible outline levels.
#
levels = [
    'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Level 6',
    'Level 7', 'Level 6', 'Level 5', 'Level 4', 'Level 3', 'Level 2',
    'Level 1']

worksheet4.write_column('A1', levels)

worksheet4.set_row(0, None, None, {'level': 1})
worksheet4.set_row(1, None, None, {'level': 2})
worksheet4.set_row(2, None, None, {'level': 3})
worksheet4.set_row(3, None, None, {'level': 4})
worksheet4.set_row(4, None, None, {'level': 5})
worksheet4.set_row(5, None, None, {'level': 6})
worksheet4.set_row(6, None, None, {'level': 7})
worksheet4.set_row(7, None, None, {'level': 6})
worksheet4.set_row(8, None, None, {'level': 5})
worksheet4.set_row(9, None, None, {'level': 4})
worksheet4.set_row(10, None, None, {'level': 3})
worksheet4.set_row(11, None, None, {'level': 2})
worksheet4.set_row(12, None, None, {'level': 1})

workbook.close()

#%%
class Excel_report():
    import xlsxwriter
    """
    Klasa dla przygotowania raportu w Excel w oparciu o Dataframe
    """
        
    def __init__(self, dataframe):
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
        
        
        
    def unload(self, path, sheet_name, groupby):
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
        import pandas as pd
        
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            row = 0
            #            iteracja po grupach
            #TODO: dodac kolorowanie komorek oraz kolejne poziomy
            df_groups = self.dataframe.groupby(by=groupby)
            for name, group in df_groups:
#                group.drop_duplicates(,inplace=True)
#                no_rows = group[group.columns[0]].count()
                
#                for n, g in df[['obreb','funkcja_dominujaca','mobile','wired']].groupby(by=['obreb','funkcja_dominujaca']):
#                    print(n)
#                    print(g.sum())
                
                df = group.sum()
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row+1, header=True)
                
                workbook  = writer.book
                worksheet = writer.sheets[sheet_name]
                worksheet.set_row(row,None,None, {'level':0})
                
#                worksheet.write(row,0,'kolumna1')
#                for xrow in range(row+1, row+1+no_rows+1):
#                    worksheet.set_row(xrow,None,None,{'level':1})
#                row = row+no_rows+1
                row = row +1
                worksheet.set_column('A:C', 10)
#            worksheet.set_column('A:C", 10)