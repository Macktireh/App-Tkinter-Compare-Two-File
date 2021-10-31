import pandas as pd
import numpy as np
import pyttsx3
import os

from datetime import date
from openpyxl import load_workbook

today = date.today()
folder_exp = f'C:/Users/J1049122/Desktop/Station Data/Master-Data/export/AFR_{today}'
os.mkdir(folder_exp)

path_data_SAP = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Data-SAP-1.xlsx"
path_data_sharepoint = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/sharepoint-AFR.xlsx"
path_list = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Affiliate_list.xlsx"



def com(df_X, df_Y, col):

    diff_X = np.setdiff1d(df_X[col] ,df_Y[col])
    ecart_X = df_X.loc[df_X[col].isin(diff_X)]
    
    print("Données SAP versus données Sharepoint :")
    print(f"il y'a {len(diff_X)} code SAP de différence")
    
    print()
    diff_Y = np.setdiff1d(df_Y[col], df_X[col])
    ecart_Y = df_Y.loc[df_Y[col].isin(diff_Y)]
    
    print("Données Sharepoint versus données SAP :")
    print(f"il y'a {len(diff_Y)} code SAP de différence")

    commun = df_X.loc[~df_X[col].isin(diff_X)]

    return ecart_X, ecart_Y, commun


def com_1(df_X, df_Y, col):

    diff_X = np.setdiff1d(df_X[col] ,df_Y[col])
    ecart_X = df_X.loc[df_X[col].isin(diff_X)]
    
    # print("Données SAP versus données Sharepoint :")
    # print(f"il y'a {len(diff_X)} code SAP de différence")
    
    # print()

    diff_Y = np.setdiff1d(df_Y[col], df_X[col])
    ecart_Y = df_Y.loc[df_Y[col].isin(diff_Y)]
    
    # print("Données Sharepoint versus données SAP :")
    # print(f"il y'a {len(diff_Y)} code SAP de différence")

    commun = df_X.loc[~df_X[col].isin(diff_X)]

    return ecart_X, ecart_Y, commun


def comparer(Pays):

    for i in Pays:

        element = i

        print()

        print('-'*20)
        print(f"Pays : {element}")
        print('-'*20)

        path_ecart = f"{folder_exp}/{element + '_' + str(today)}.xlsx"


        df_sap = pd.read_excel(path_data_SAP, sheet_name=element)
        df_sap.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
        df_sap = df_sap.drop_duplicates(subset = "SAPCode", keep = 'first')
        dim_sap = df_sap.shape
        print(f"dimension données SAP pour {element} est : {dim_sap}")
        df_sap['SAPCode'] = df_sap['SAPCode'].str.strip()

        df_sharepoint = pd.read_excel(path_data_sharepoint, sheet_name=element)
        df_sharepoint = df_sharepoint.drop_duplicates()
        dim_sharepoint = df_sharepoint.shape
        print(f"dimension données sharepoint pour {element} est : {dim_sharepoint}")
        df_sharepoint['SAPCode'] = df_sharepoint['SAPCode'].str.strip()

        print()

        print("Comparaison :")
        print('-'*7)

        X, Y, df_commun_1 = com(df_sap, df_sharepoint, 'SAPCode')
        a, cost, df_commun_2 = com_1(df_commun_1, df_sharepoint, 'SAPCode_BM')
        b, cost, df_commun_3 = com_1(df_commun_2, df_sharepoint, 'SAPCode_BM_ISACTIVESITE')

        writer = pd.ExcelWriter(path_ecart, engine = 'openpyxl')
        df_sap.to_excel(writer, sheet_name = 'Data_SAP_Brute', index=False)
        df_sharepoint.to_excel(writer, sheet_name = 'Data_Sharepoint_Brute', index=False)
        X.to_excel(writer, sheet_name = 'ecart_SAP_vs_Sharepoint', index=False)
        Y.to_excel(writer, sheet_name = 'ecart_Sharepoint_vs_SAP', index=False)
        a.to_excel(writer, sheet_name = 'SAP_vs_Sharepoint_SAPCode_BM', index=False)
        b.to_excel(writer, sheet_name = 'SAP_vs_Sharepoint_SAPCode_BM_ISACTIVESITE', index=False)

        writer.save()
        writer.close()

        #print()
        print('#'*100)
        print()


        # sh = pd.read_excel("C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Data-SAP.xlsx")
        # sh = sh.drop_duplicates()
        # sh['SAPCode'] = sh['SAPCode'].str.strip()

        # z = sh['Affiliate'].unique()

        # # for w in z:
        # d = sh[sh['Affiliate']==w]

        sh = df_sharepoint.copy()

        if a.shape[0] > 0:
            for j in range(a.shape[0]):
                for k in range(sh.shape[0]):
                    if a['SAPCode'].iloc[j] == sh['SAPCode'].iloc[k]:
                        sh['BUSINESSMODEL'].iloc[k] = a['BUSINESSMODEL'].iloc[j]
                        sh['BM_source'].iloc[k] = a['BM_source'].iloc[j]


        book = load_workbook(path_list)
        writer_list = pd.ExcelWriter(path_list, engine = 'openpyxl')
        writer_list.book = book

        

        sh.to_excel(writer_list, sheet_name = element, index=False)
        writer_list.save()
        writer_list.close()







pays = ['Botswana', 'Ghana', 'Kenya', 'Mauritius', 'Malawi', 'Mozambique', 'Namibia',
 'Nigeria', 'Tanzania', 'Uganda', 'South Africa', 'Zambia',
 'Zimbabwe', 'Central Afr.Rep', 'Congo', 'Cameroon', 'Gabon', 'Guinea Conakry',
 'Equatorial Gui.', 'Morocco', 'Mali', 'Senegal', 'Chad', 'Togo', 'Mayotte']

comparer(pays)


print()
print("--------------------")
print("Terminer avec succès")
print("--------------------")
print()
