import pandas as pd
import numpy as np
import pyttsx3


def import_export(pays):

    path_data_SAP = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/SAP-P2K-1.xlsx"
    path_data_sharepoint = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/sharepoint-AFR.xlsx"
    path_ecart = f"C:/Users/J1049122/Desktop/Station Data/Master-Data/ecart/{pays}.xlsx"

    df_sharepoint = pd.read_excel(path_data_sharepoint, sheet_name=pays)
    df_sap = pd.read_excel(path_data_SAP, sheet_name=pays)
    df_ecart_1 = pd.read_excel(path_ecart, sheet_name= "ecart_SAP_vs_Sharepoint_"+pays)
    df_ecart_2 = pd.read_excel(path_ecart, sheet_name= "ecart_Sharepoint_vs_SAP_"+pays)
    df_en_commun = pd.read_excel(path_ecart, sheet_name= "en_commun_"+pays)

    path_export = f"C:/Users/J1049122/Desktop/Station Data/Master-Data/exp/{pays}.xlsx"

    writer = pd.ExcelWriter(path_export, engine = 'openpyxl')

    df_sharepoint.to_excel(writer, sheet_name = 'Data_Sharepoint_Brute_'+pays, index=False)
    df_sap.to_excel(writer, sheet_name = 'Data_SAP_Brute_'+pays, index=False)
    df_ecart_1.to_excel(writer, sheet_name = "ecart_SAP_vs_Sharepoint_"+pays, index=False)
    df_ecart_2.to_excel(writer, sheet_name = "ecart_Sharepoint_vs_SAP_"+pays, index=False)
    df_en_commun.to_excel(writer, sheet_name = "en_commun_"+pays, index=False)

    writer.save()
    writer.close()


Pays = ['Botswana', 'Ghana', 'Nigeria', 'Tanzania', 'Kenya', 'Mauritius', 'Malawi', 'Mozambique', 'Namibia', 'Uganda', 'South Africa', 'Zambia', 'Zimbabwe']

for pays in Pays:
    import_export(pays)



print()
print("--------------------")
print("Terminer avec succès")
print("--------------------")
print()