import pandas as pd
import numpy as np
import pyttsx3
import os
import shutil

from datetime import date
from openpyxl import load_workbook

today = date.today()

path_import = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Data-SAP.xlsx"
path_export = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Data-SAP-1.xlsx"

# def sap(env):
#     if len(env) > 1:
#         data = []
#         if len(env) == 2:
#             for i in env:
#                 data.append(pd.read_excel(path_import[i]))
#             return data[0], data[1]

#         elif len(env) == 3:
#             for i in env:
#                 data.append(pd.read_excel(path_import[i]))
#             return data[0], data[1], data[3]
        
#     else:
#         return pd.read_excel(path_import[env[0]])


# if os.path.exists(path_export):
#     os.remove(path_export)
#     print(f"le fichier {path_export} à été bien supprimer\n-------------")
# else:
#     print(f"Impossible de supprimer le fichier {path_export} car il n'existe pas\n-------------")

df = pd.read_excel(path_import)
print('All Data SAP', df.shape)
# print(df.head())
pays = df['Affiliate'].unique()
print(pays)

for sheet in pays:
    book = load_workbook(path_export)
    writer = pd.ExcelWriter(path_export, engine = 'openpyxl')
    writer.book = book

    d = df[df['Affiliate']==sheet]

    d.to_excel(writer, sheet_name = sheet, index=False)
    writer.save()
    writer.close()

print()
print("--------------------")
print("Terminer avec succès")
print("--------------------")
print()