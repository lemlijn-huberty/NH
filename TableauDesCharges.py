import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

# Version 25/12/2020.
# only > XLSX < files.

months = [("janvier"), ("février"), ("mars"), ("avril"), ("mai"), ("juin"), ("juillet"), ("août"), ("septembre"), ("octobre"), ("novembre"), ("décembre")]

files = np.array(['Antoine','CHA','CSL','CUP','Chanly','HP','IFAC','LaBouv','MSP','Séniori','Vielsal'])

def getIndexes(dfObj, value):
    listOfPos = []
    result = dfObj.isin([value])
    seriesObj = result.any()
    columnNames = list(seriesObj[seriesObj == True].index)
    for col in columnNames:
        rows = list(result[col][result[col] == True].index)
        for row in rows:
            listOfPos.append((row, col))
    return listOfPos

def execute():
   excel = str(root.filename).replace("/", "\\")
   df = pd.read_excel(excel, engine='openpyxl')

   for file in files:
       lookup = 'VIVALIA (Vivalia_' + file + ')'
       ListOfPositions = getIndexes(df, lookup)
       df2 = pd.DataFrame(ListOfPositions, columns=['idx', 'col'])
       val_from = df2['idx'].iloc[0]
       val_to = val_from + 85
       df_target = df[val_from:val_to]
       df_target.columns = ["", "", "", "Total 2020", "janvier", "février", "mars", "avril", "mai", "juin", "juillet",
                            "août", "septembre", "octobre", "novembre", "décembre"]
       target_excel = str(excel).replace(".xlsx", "_") + file + '.xlsx'
       print(lookup + '_' + str(val_from) + '-' + str(val_to) + ' ' + target_excel)
       writer = pd.ExcelWriter(target_excel, engine='xlsxwriter')
       df_target.to_excel(writer, sheet_name='Liste', index=False)
       workbook = writer.book
       workbook.formats[0].set_font_size(8)
       workbook.formats[0].set_font_name('Arial')
       worksheet = writer.sheets['Liste']
       format_num = workbook.add_format({'num_format': '#,##0.00', 'font_size': '8'})
       format_num.set_font_size(8)
       format_num.set_font_name('Arial')
       worksheet.set_column('D:P', 14, format_num)
       worksheet.set_column('C:C', 35, None)
       format_bold = workbook.add_format({'bold': True, 'font_color': 'black', 'num_format': '#,##0.00'})
       format_bold.set_font_size(9)
       format_bold.set_font_name('Arial')
       format_total = workbook.add_format({'bold': True, 'font_color': 'blue', 'num_format': '#,##0.00'})
       format_total.set_font_size(9)
       format_total.set_font_name('Arial')

       df_sum = pd.read_excel("c:\Python\Tableau des charges - formules totaux.xlsx", engine='openpyxl')
       df_row_tot = pd.read_excel("C:\Python\Tableau des charges - lignes total.xlsx", engine='openpyxl')

       columns = np.array(['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'])

       for ind in df_sum.index:
           row_idx = str(df_sum['r'][ind])
           for column in columns:
               cell = str(column + row_idx)
               formula = str('=SUM(' + column + str(df_sum['f'][ind]) + ':' + column + str(df_sum['t'][ind]) + ')')
               worksheet.write_formula(cell, formula)
           worksheet.set_row(df_sum['r'][ind] - 1, None, format_bold)

       for ind in df_row_tot.index:
           row_idx = str(df_row_tot['r'][ind])
           cell = str('D' + row_idx)
           cell_from = str(columns[0] + row_idx)
           cell_to = str(columns[-1] + row_idx)
           #print(cell_from + ':' + cell_to)
           formula = str('=SUM(' + str(cell_from) + ':' + str(cell_to) + ')')
           worksheet.write_formula(cell, formula)

           idx_0_col = g_month.get() + 1
           while idx_0_col < len(columns):
             cell = columns[idx_0_col] + row_idx
             worksheet.write_formula(cell, "=0")
             idx_0_col += 1

       for col in columns:
           cell = str(col + '69')
           formula = str('=' + col + '5+' + col + '32+' + col + '46+' + col + '59+' + col + '64')
           #print(cell + ' ' + formula)
           worksheet.write_formula(cell, formula)
           cell = str(col + '82')
           formula = str('=' + col + '72+' + col + '77')
           worksheet.write_formula(cell, formula)
           cell = str(col + '84')
           formula = str('=' + col + '82+' + col + '69')
           worksheet.write_formula(cell, formula)

       worksheet.set_row(69 - 1, None, format_total)
       worksheet.set_row(82 - 1, None, format_total)
       worksheet.set_row(84 - 1, None, format_total)

       cell_to = str(columns[g_month.get()] + '86')
       formula = str('=AVERAGE(E86:' + str(cell_to) + ')')
       worksheet.write_formula('D86', formula)
       idx_0_col = g_month.get() + 1
       while idx_0_col < len(columns):
           cell = columns[idx_0_col] + '86'
           worksheet.write_formula(cell, "=0")
           idx_0_col += 1

       writer.save()
   quit()

root = tk.Tk()
g_month = tk.IntVar()
tk.Label(root, text="""Choisir le mois pour les totaux:""", justify = tk.LEFT, padx = 20).pack()
for val, month in enumerate(months):
    tk.Radiobutton(root, text=month, padx = 20, variable=g_month, command=execute, value=val).pack(anchor=tk.W)

root.geometry("400x350+30+30")
root.title("Traitement - Tableau des charges")
root.filename = filedialog.askopenfilename(initialdir = "c:/",title = "Tableau des charges",filetypes = (("excel files","*.xls*"),("all files","*.*")))
root.mainloop()
