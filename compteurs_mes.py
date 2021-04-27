#NioulBoy 2021

import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import datetime
"""
Prepare excel file to have all activated meters and connexion type
"""
#r"C:\Users\Mandiaye\Documents\Final 2020\Compteurs2020\Compteurs MES2020.xlsx"
print("Entrer le path du fichier de compensation")
path = input()
df = pd.ExcelFile(path)
print(df.sheet_names)

"""
Identify sheets and use it as df
"""
df1 = pd.read_excel(df, sheet_name= 'Feuil1')
print(df1.dtypes)
print(df1.keys())
print(df1)

# dropping ALL duplicte values 
df1.drop_duplicates(subset ="N° de compteur", 
                     keep = 'first', inplace = True) 
print(df1)

#TODO 'find type de connexion' using pandas


#Create a Pandas Excel writer using openpyxl as the engine.
writer = pd.ExcelWriter(path, engine='openpyxl')
# try to open an existing workbook
writer.book = load_workbook(path)
# copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
#Sort by the values along either axis.
df1.sort_values(by=["Date d'activation"]).to_excel(writer, sheet_name='Feuil1',startcol=10)


# Close the Pandas Excel writer and output the Excel file.
writer.save()


writer.close()
