#NioulBoy

import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt

#Juillet2020.xlsx
#r"C:\Users\Mandiaye\Documents\Final 2020\Juillet2020\Fichier Compensation SCL Juillet2020.xlsx"
#C:\Users\Mandiaye\Documents\Final 2020\Juillet2020\Fichier Compensation SCL Juillet2020.xlsx

"""
Load in the Excel data that represents a month state for our SCL Energie Solutions company.
"""
print("Entrer le path du fichier de compensation")
path = input()
df = pd.ExcelFile(path)
print(df.sheet_names)

"""
Identify sheets and use it as df
"""
df1 = pd.read_excel(df, sheet_name= '1-Liste clients ')
df2 = pd.read_excel(df, sheet_name='3-Nouveaux Clients')
df3 = pd.read_excel(df,sheet_name='5-Client en Inst de MES')
df4 = pd.read_excel(df,sheet_name='7-Liste RO')
print(df1.dtypes)ddffd
print(df1.keys())
print(df4.keys())
print("********************************")

"""
Count customers by village
"""
print("************* Clients par villages*******************")
village = df1.groupby(['Village']).count()
recordVillage = village['Etats']
print(recordVillage)
sumVillage = recordVillage.sum()
print(sumVillage)
print("********************************")
print(recordVillage.describe())

"""
Count customers by services
"""
print("************clients par services********************")
usage = df1.groupby(['Type du niveau de service ']).count()
print(usage['Etats'])
print("---------------------------------------")

"""
Count customers by services
items with 'Type du niveau de service' that start with S4.
count 'Usage d'utilisation'
"""
print("**************payant domestique et productif S4******************")
map = df1[df1["Type du niveau de service "].map(lambda x: x.startswith('S4'))]
utilisation = map.groupby(["Usage d'utilisation"]).count()
print(utilisation["Etats"])
print("********************************")

"""
Count new customers by services 
"""
print("********************Nouveaux Clients******************")
print("---------------------------------------")
print("************Nouveau clients par services********************")
usageNew = df2.groupby(['Type du niveau de service ']).count()
print(usageNew['Etats'])
print("---------------------------------------")

"""
Count new customers by services
items with 'Type du niveau de service' that start with S4.
count 'Usage d'utilisation'
"""
print("**************Nouveau payant domestique et productif S4******************")
mapNew = df2[df2["Type du niveau de service "].map(lambda x: x.startswith('S4'))]
utilisationNew = mapNew.groupby(["Usage d'utilisation"]).count()
print(utilisationNew["Etats"])
print("********************************")

"""
Pending customers by services
"""
print("************Clients en instance********************")
clientsEnInstance = df3.groupby(['Type du niveau de service ']).count()
print(clientsEnInstance['Statuts'])
print("********************************")

"""
RO charge by services
"""
print("************Nombre RO********************")
nombreROparService = df4.groupby(['Type client']).count()
print(nombreROparService['Mois de recharge'])

"""
Plot some figures
"""
plt.figure()
recordVillage.plot(title= 'Clients par village', color= 'purple', kind = 'bar')
plt.show(block=False)
plt.figure()
usage['Etats'].plot(title= 'Nombre Client par Service', color= 'green', kind = 'bar', y="usage['Etats']", rot = 0)
plt.show(block=False)
plt.figure()
nombreROparService['Mois de recharge'].plot(title= 'Nombre RO par Service', kind = 'bar', rot = 0)
plt.show()

"""
Write results on the excel file
"""
# Create a Pandas Excel writer using openpyxl as the engine.
writer = pd.ExcelWriter(path, engine='openpyxl')
# try to open an existing workbook
writer.book = load_workbook(path)
# copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

#Clients
recordVillage.to_excel(writer, sheet_name='2-Client par Vil - Serv',startcol=7)
usage['Etats'].to_excel(writer,sheet_name='2-Client par Vil - Serv',startcol=7,startrow=len(recordVillage)+3)
utilisation["Etats"].to_excel(writer,sheet_name='2-Client par Vil - Serv',startcol=7,startrow=len(usage['Etats'])+len(recordVillage)+3)
#Nouveaux Clients
usageNew['Etats'].to_excel(writer,sheet_name='4-Nouv Clients par Serv', startcol= 8)
utilisationNew['Etats'].to_excel(writer,sheet_name='4-Nouv Clients par Serv',startrow=len(usageNew['Etats']), startcol=8)
#Clients en instance
clientsEnInstance['Statuts'].to_excel(writer,sheet_name='6-Client en inst par Serv',startcol=6)
#Nombre RO
nombreROparService['Mois de recharge'].to_excel(writer,sheet_name='8-Nombre RO',startcol=8)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
writer.close()