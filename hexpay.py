#NioulBoy 2021

import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import datetime
"""
Prepare excel file to have daily transaction from Hexpay
"""

print("Entrer le path du fichier de compensation")
#test path
#path = r"C:\Users\Mandiaye\Documents\Final 2021\Mai2021\RAPPORT MENSUEL SCL MAI 2021.xlsx"
path = input()
df = pd.ExcelFile(path)
print(df.sheet_names)

"""
Identify sheets and use it as df. save as Excel and delete tests
"""
df1 = pd.read_excel(df, sheet_name= 'Sheet0')
print(df1.dtypes)
print(df1.keys())

recharge_journalier = df1.groupby(['Transaction Date'])
liste = recharge_journalier['Purchase Amount（XOF）'].sum()
print(liste)

"""
Write results on the excel file
"""
# Create a Pandas Excel writer using openpyxl as the engine.
writer = pd.ExcelWriter(path, engine='openpyxl')
# try to open an existing workbook
writer.book = load_workbook(path)
# copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

liste.to_excel(writer, sheet_name='Sheet0',startcol=7)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
writer.close()