# importer la binliothèque
import tkinter as tk
import pandas as pd
import numpy as np
import pyttsx3
import os
import shutil
import time

from tkinter import filedialog, messagebox, ttk
from tkinter.constants import ACTIVE
from datetime import date
from openpyxl import load_workbook


############################################################################################
# --------------------------------------  Frontend  -------------------------------------- #
############################################################################################

# création de l'objet de la fenetre
root = tk.Tk()

# personnaliser la fenetre
root.title("    PyApp Station Data Desktop")  # nom d'entête de la fenetre
root.iconbitmap("TotalEnergies.ico")  # icone de la fenetre
root.geometry("900x600+15+15")  # taille de la fenetre
root.minsize(900, 600)
root.maxsize(1000, 700)

# configuration du font de la fenetre (couleur ou autre)
# root.config(background='#CCCCCC')

# barre de menu
mainMenu = tk.Menu(root)

file_menu = tk.Menu(root, tearoff=0)
file_menu.add_command(label="A propos")
file_menu.add_command(label="Quit", command=root.quit)

mainMenu.add_cascade(label="File", menu=file_menu)


def compare():
    """Si le fichier sélectionné est valide, cela chargera le fichier"""

    # file 1
    file_path_1 = label_file_1["text"]
    try:
        excel_filename = r"{}".format(file_path_1)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename)
        else:
            if var_entry_1.get() == "":
                df1 = pd.read_excel(excel_filename)
            else:
                df1 = pd.read_excel(
                    excel_filename, sheet_name=var_entry_1.get())
    except ValueError:
        tk.messagebox.showerror(
            "Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror(
            "Information", f"No such file as {file_path_1}")
        return None

    # file 2
    file_path_2 = label_file_2["text"]
    try:
        excel_filename = r"{}".format(file_path_2)
        if excel_filename[-4:] == ".csv":
            df2 = pd.read_csv(excel_filename)
        else:
            if var_entry_2.get() == "":
                df2 = pd.read_excel(excel_filename)
            else:
                df2 = pd.read_excel(
                    excel_filename, sheet_name=var_entry_2.get())
    except ValueError:
        tk.messagebox.showerror(
            "Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror(
            "Information", f"No such file as {file_path_2}")
        return None

    today = date.today()

    folder_result = "{}/resutl_{}".format(lbl1["text"], today)

    if os.path.exists(folder_result):
        shutil.rmtree(f'{folder_result}')
        print(
            f"le dossier {folder_result} à été bien supprimer et recréer\n-------------")
        print()
    else:
        print(f"le dossier {folder_result} n'existe pas\n-------------")
        print()

    os.mkdir(folder_result)

    folder_exp = f'{folder_result}/testAFR_{today}'

    if os.path.exists(folder_exp):
        shutil.rmtree(f'{folder_exp}')
        print(
            f"le dossier AFR_{today} à été bien supprimer et recréer\n-------------")
        print()
    else:
        print(f"le dossier AFR_{today} n'existe pas\n-------------")
        print()

    os.mkdir(folder_exp)

    data_sap = df1.copy()
    data_sharepoint = df2.copy()

    sh_p = data_sharepoint['Affiliate'].unique()
    sap_p = data_sap['Affiliate'].unique()

    for i in sh_p:

        if i in sap_p:

            element = i

            print()

            print('-'*20)
            print(f"Pays : {element}")
            print('-'*20)

            path_ecart = f"{folder_exp}/{element + '_' + str(today)}.xlsx"
            #path_list = f"{folder_list_affiliate}/list_affiliate_{str(today)}.xlsx"

            df_sap = data_sap[data_sap['Affiliate'] == element]
            df_sap.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
            df_sap = df_sap.drop_duplicates(subset="SAPCode", keep='first')
            dim_sap = df_sap.shape
            print(f"dimension données SAP pour {element} est : {dim_sap}")
            df_sap['SAPCode'] = df_sap['SAPCode'].str.strip()

            df_sharepoint = data_sharepoint[data_sharepoint['Affiliate'] == element]
            df_sharepoint = df_sharepoint.drop_duplicates()
            dim_sharepoint = df_sharepoint.shape
            print(
                f"dimension données sharepoint pour {element} est : {dim_sharepoint}")
            df_sharepoint['SAPCode'] = df_sharepoint['SAPCode'].str.strip()

            print()

            print("Comparaison :")
            print('-'*7)

            X, Y, df_commun_1 = com(
                df_sap, df_sharepoint, col_name_1["text"], col_name_2["text"])

            writer = pd.ExcelWriter(path_ecart, engine='openpyxl')
            df_sap.to_excel(writer, sheet_name='Data_SAP_Brute', index=False)
            df_sharepoint.to_excel(
                writer, sheet_name='Data_Sharepoint_Brute', index=False)
            X.to_excel(
                writer, sheet_name='ecart_SAP_vs_Sharepoint', index=False)
            Y.to_excel(
                writer, sheet_name='ecart_Sharepoint_vs_SAP', index=False)

            writer.save()
            writer.close()


def selected_item_1():
    for i in box1.curselection():
        # var_col_name_1.set(box1.get(i))
        col_name_1["text"] = box1.get(i)


def selected_item_2():
    for j in box2.curselection():
        # var_col_name_2.set(box2.get(i))
        col_name_2["text"] = box2.get(j)


def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    lbl1["text"] = filename


# ---------- la boîte de dialogue d'ouverture de fichier ---------- #
file_frame_1 = tk.LabelFrame(
    root, text="Open First File", background='#CCCCCC')
file_frame_1.place(height=200, width=400, rely=0.05, relx=0.02)

# label
label_1 = tk.Label(
    file_frame_1, text='If the file is an Excel file enter the name of the sheet (optional)')
label_1.place(rely=0.45, relx=0)

var_entry_1 = tk.StringVar()
sheet_name_1 = tk.Entry(file_frame_1, textvariable=var_entry_1)
sheet_name_1.place(rely=0.65, relx=0.10)

# Buttons
button1 = tk.Button(file_frame_1, text="Browse A File",
                    command=lambda: File_dialog_1())
button1.place(rely=0.85, relx=0.50)

button2 = tk.Button(file_frame_1, text="Load File",
                    command=lambda: view_data())
button2.place(rely=0.85, relx=0.30)

# Le texte du fichier/chemin d'accès au fichier
label_file_1 = ttk.Label(file_frame_1, text="No File Selected")
label_file_1.place(rely=0, relx=0)

box1 = tk.Listbox(root)
box1.place(height=200, width=200, rely=0.43, relx=0.05)

# commande signifie mettre à jour la vue de l'axe y du widget
treescrolly = tk.Scrollbar(box1, orient="vertical", command=box1.yview)
# commande signifie mettre à jour la vue axe x du widget
treescrollx = tk.Scrollbar(box1, orient="horizontal", command=box1.xview)
# affecter les barres de défilement au widget Treeview
box1.configure(xscrollcommand=treescrollx.set,
               yscrollcommand=treescrolly.set)
# faire en sorte que la barre de défilement remplisse l'axe x du widget Treeview
treescrollx.pack(side="bottom", fill="x")
# faire en sorte que la barre de défilement remplisse l'axe y du widget Treeview
treescrolly.pack(side="right", fill="y")

# colonne selectionner
# var_col_name_1 = tk.StringVar()
# var_col_name_1.trace("w", selected_item_1)
# col_name_1 = tk.Label(root, textvariable=var_col_name_1,
#                       background="#BFF3EC", width=22)
col_name_1 = tk.Label(root, text="",
                      background="#74BBE4", width=22)
col_name_1.place(rely=0.53, relx=0.28)

btn_1 = tk.Button(root, text='Ok', command=selected_item_1)
btn_1.place(rely=0.63, relx=0.30)

# ---------- la boîte de dialogue d'ouverture de fichier ---------- #
file_frame_2 = tk.LabelFrame(
    root, text="Open Second File", background='#CCCCCC')
file_frame_2.place(height=200, width=400, rely=0.05, relx=0.50)

# label
label_2 = tk.Label(
    file_frame_2, text='If the file is an Excel file enter the name of the sheet (optional)')
label_2.place(rely=0.45, relx=0)

var_entry_2 = tk.StringVar()
sheet_name_2 = tk.Entry(file_frame_2, textvariable=var_entry_2)
sheet_name_2.place(rely=0.65, relx=0.10)

# Buttons
button3 = tk.Button(file_frame_2, text="Browse A File",
                    command=lambda: File_dialog_2())
button3.place(rely=0.85, relx=0.50)

button4 = tk.Button(file_frame_2, text="Load File",
                    command=lambda: view_data_2())
button4.place(rely=0.85, relx=0.30)

# Le texte du fichier/chemin d'accès au fichier
label_file_2 = ttk.Label(file_frame_2, text="No File Selected")
label_file_2.place(rely=0, relx=0)


box2 = tk.Listbox(root)
box2.place(height=200, width=200, rely=0.43, relx=0.53)

# commande signifie mettre à jour la vue de l'axe y du widget
treescrollw = tk.Scrollbar(box2, orient="vertical", command=box2.yview)
# commande signifie mettre à jour la vue axe x du widget
treescrollz = tk.Scrollbar(box2, orient="horizontal", command=box2.xview)
# affecter les barres de défilement au widget Treeview
box2.configure(xscrollcommand=treescrollz.set,
               yscrollcommand=treescrollw.set)
# faire en sorte que la barre de défilement remplisse l'axe x du widget Treeview
treescrollz.pack(side="bottom", fill="x")
# faire en sorte que la barre de défilement remplisse l'axe y du widget Treeview
treescrollw.pack(side="right", fill="y")

# colonne selectionner
# var_col_name_2 = tk.StringVar()
# var_col_name_2.trace("w", selected_item_2)
# col_name_2 = tk.Label(root, textvariable=var_col_name_2,
#   background="#BFF3EC", width=22)
col_name_2 = tk.Label(root, text="",
                      background="#74BBE4", width=22)
col_name_2.place(rely=0.53, relx=0.76)

btn_2 = tk.Button(root, text='Ok', command=selected_item_2)
btn_2.place(rely=0.63, relx=0.80)


button_comparer = tk.Button(
    root, text="Compare", width=20, background="#004C8C", fg="white", command=compare)
button_comparer.place(rely=0.90, relx=0.30)

button_quit = tk.Button(
    root, text="Quit", width=20, background="#C60030", fg="white", command=root.quit)
button_quit.place(rely=0.90, relx=0.50)


fram = tk.Frame(root, bd=1)
# folder_path = tk.StringVar()
# lbl1 = tk.Label(fram, textvariable=folder_path)
lbl1 = tk.Label(fram, text="")
lbl1.grid(row=0, column=0)
button_folder = tk.Button(
    fram, text="destination folder", command=browse_button)
button_folder.grid(row=1, column=0)

fram.place(rely=0.80, relx=0.25)


# button_comparer = tk.Button(
#     root, text="compare", width=20, background="#3FB8F2", command=lambda: compare)
# button_folder.place(rely=0.90, relx=0.40)

###########################################################################################
# --------------------------------------  Backend  -------------------------------------- #
###########################################################################################


def File_dialog_1():
    """Cette fonction ouvrira l'explorateur de fichiers et affectera le chemin de fichier choisi à label_file"""
    filename_1 = filedialog.askopenfilename(initialdir="E:\Total\Station Data\Master data\Data source",
                                            title="Select A File",
                                            filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file_1["text"] = filename_1
    return None


def view_data():
    new_interface = tk.Toplevel(root)
    new_interface.title("Previous Data of first file")
    new_interface.iconbitmap("TotalEnergies.ico")
    new_interface.geometry("800x550")
    new_interface.resizable(width=False, height=False)

    frame1 = tk.LabelFrame(new_interface, text="Excel Data")
    frame1.place(height=500, width=750, rely=0.05, relx=0.05)

    tv1 = ttk.Treeview(frame1)
    tv1.place(relheight=1, relwidth=1)

    # commande signifie mettre à jour la vue de l'axe y du widget
    treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)

    # commande signifie mettre à jour la vue axe x du widget
    treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)

    # affecter les barres de défilement au widget Treeview
    tv1.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)

    # faire en sorte que la barre de défilement remplisse l'axe x du widget Treeview
    treescrollx.pack(side="bottom", fill="x")

    # faire en sorte que la barre de défilement remplisse l'axe y du widget Treeview
    treescrolly.pack(side="right", fill="y")

    def Load_excel_data_1():
        """Si le fichier sélectionné est valide, cela chargera le fichier"""
        file_path_1 = label_file_1["text"]
        try:
            excel_filename = r"{}".format(file_path_1)
            if excel_filename[-4:] == ".csv":
                df1 = pd.read_csv(excel_filename)
                for id, column in enumerate(df1.columns):
                    box1.insert(id, column)
            else:
                if var_entry_1.get() == "":
                    df1 = pd.read_excel(excel_filename)
                    for id, column in enumerate(df1.columns):
                        box1.insert(id, column)
                else:
                    df1 = pd.read_excel(
                        excel_filename, sheet_name=var_entry_1.get())
                    for id, column in enumerate(df1.columns):
                        box1.insert(id, column)

        except ValueError:
            tk.messagebox.showerror(
                "Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            tk.messagebox.showerror(
                "Information", f"No such file as {file_path_1}")
            return None

        clear_data()
        tv1["column"] = list(df1.columns)
        tv1["show"] = "headings"
        for column in tv1["columns"]:
            tv1.heading(column, text=column)

        df_rows = df1.to_numpy().tolist()
        for row in df_rows:
            tv1.insert("", "end", values=row)

        return df1

    def clear_data():
        tv1.delete(*tv1.get_children())
        return None

    Load_excel_data_1()


def File_dialog_2():
    """Cette fonction ouvrira l'explorateur de fichiers et affectera le chemin de fichier choisi à label_file"""
    filename_2 = filedialog.askopenfilename(initialdir="E:\Total\Station Data\Master data\Data source",
                                            title="Select A File",
                                            filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file_2["text"] = filename_2
    return None


def view_data_2():
    new_interface = tk.Toplevel(root)
    new_interface.title("Previous Data of second file")
    new_interface.iconbitmap("TotalEnergies.ico")
    new_interface.geometry("800x550")
    new_interface.resizable(width=False, height=False)

    frame2 = tk.LabelFrame(new_interface, text="Excel Data")
    frame2.place(height=500, width=750, rely=0.05, relx=0.05)

    tv2 = ttk.Treeview(frame2)
    tv2.place(relheight=1, relwidth=1)

    # commande signifie mettre à jour la vue de l'axe y du widget
    treescrollw = tk.Scrollbar(frame2, orient="vertical", command=tv2.yview)

    # commande signifie mettre à jour la vue axe x du widget
    treescrollz = tk.Scrollbar(frame2, orient="horizontal", command=tv2.xview)

    # affecter les barres de défilement au widget Treeview
    tv2.configure(xscrollcommand=treescrollz.set,
                  yscrollcommand=treescrollw.set)

    # faire en sorte que la barre de défilement remplisse l'axe x du widget Treeview
    treescrollz.pack(side="bottom", fill="x")

    # faire en sorte que la barre de défilement remplisse l'axe y du widget Treeview
    treescrollw.pack(side="right", fill="y")

    def Load_excel_data_2():
        """Si le fichier sélectionné est valide, cela chargera le fichier"""
        file_path_2 = label_file_2["text"]
        try:
            excel_filename = r"{}".format(file_path_2)
            if excel_filename[-4:] == ".csv":
                df2 = pd.read_csv(excel_filename)
                for id, column in enumerate(df2.columns):
                    box2.insert(id, column)
            else:
                if var_entry_2.get() == "":
                    df2 = pd.read_excel(excel_filename)
                    for id, column in enumerate(df2.columns):
                        box2.insert(id, column)
                else:
                    df2 = pd.read_excel(
                        excel_filename, sheet_name=var_entry_2.get())
                    for id, column in enumerate(df2.columns):
                        box2.insert(id, column)
        except ValueError:
            tk.messagebox.showerror(
                "Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            tk.messagebox.showerror(
                "Information", f"No such file as {file_path_2}")
            return None

        clear_data()
        tv2["column"] = list(df2.columns)
        tv2["show"] = "headings"
        for column in tv2["columns"]:
            tv2.heading(column, text=column)

        df_rows = df2.to_numpy().tolist()
        for row in df_rows:
            tv2.insert("", "end", values=row)

        return df2

    def clear_data():
        tv2.delete(*tv2.get_children())
        return None

    Load_excel_data_2()


def com(df_X, df_Y, col_x, col_y, texte=True):

    if texte:
        diff_X = np.setdiff1d(df_X[col_x], df_Y[col_y])
        ecart_X = df_X.loc[df_X[col_x].isin(diff_X)]

        print("Données SAP versus données Sharepoint :")
        print(f"il y'a {len(diff_X)} code SAP de différence")

        print()
        diff_Y = np.setdiff1d(df_Y[col_y], df_X[col_x])
        ecart_Y = df_Y.loc[df_Y[col_y].isin(diff_Y)]

        print("Données Sharepoint versus données SAP :")
        print(f"il y'a {len(diff_Y)} code SAP de différence")

        commun = df_X.loc[~df_X[col_x].isin(diff_X)]

        return ecart_X, ecart_Y, commun

    else:
        diff_X = np.setdiff1d(df_X[col_x], df_Y[col_y])
        ecart_X = df_X.loc[df_X[col_x].isin(diff_X)]

        diff_Y = np.setdiff1d(df_Y[col_y], df_X[col_x])
        ecart_Y = df_Y.loc[df_Y[col_y].isin(diff_Y)]

        commun = df_X.loc[~df_X[col_x].isin(diff_X)]

        return ecart_X, ecart_Y, commun


def compare():
    """Si le fichier sélectionné est valide, cela chargera le fichier"""

    # file 1
    file_path_1 = label_file_1["text"]
    try:
        excel_filename = r"{}".format(file_path_1)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename)
        else:
            if var_entry_1.get() == "":
                df1 = pd.read_excel(excel_filename)
            else:
                df1 = pd.read_excel(
                    excel_filename, sheet_name=var_entry_1.get())
    except ValueError:
        tk.messagebox.showerror(
            "Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror(
            "Information", f"No such file as {file_path_1}")
        return None

    # file 2
    file_path_2 = label_file_2["text"]
    try:
        excel_filename = r"{}".format(file_path_2)
        if excel_filename[-4:] == ".csv":
            df2 = pd.read_csv(excel_filename)
        else:
            if var_entry_2.get() == "":
                df2 = pd.read_excel(excel_filename)
            else:
                df2 = pd.read_excel(
                    excel_filename, sheet_name=var_entry_2.get())
    except ValueError:
        tk.messagebox.showerror(
            "Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror(
            "Information", f"No such file as {file_path_2}")
        return None

    today = date.today()

    folder_result = "{}/resutl_{}".format(lbl1["text"], today)

    folder_exp = f'{folder_result}/testAFR_{today}'

    if os.path.exists(folder_exp):
        shutil.rmtree(f'{folder_exp}')
        print(
            f"le dossier AFR_{today} à été bien supprimer et recréer\n-------------")
        print()
    else:
        print(f"le dossier AFR_{today} n'existe pas\n-------------")
        print()

    os.mkdir(folder_exp)

    # folder_list_affiliate= f'E:/Total/Station Data/Master Data/export/list_affiliate_{today}'
    # os.mkdir(folder_list_affiliate)

    # path_data_SAP = "E:/Total/Station Data/Master Data/Data source/Data-SAP.xlsx"
    # path_data_sharepoint = "E:/Total/Station Data/Master Data/Data source/all-data-sharepoint.xlsx"
    # path_list = f"{folder_result}/Affiliate_list.xlsx"

    # if os.path.exists(path_list):
    #     os.remove(path_list)
    #     print("le fichier 'Affiliate_list.xlsx' à été bien supprimer et recréer\n-------------")
    # else:
    #     print("le fichier 'Affiliate_list.xlsx' n'existe pas\n-------------")

    # # data_sharepoint = pd.read_excel(
    # #     'E:/Total/Station Data/Master Data/Data source/all-data-sharepoint.xlsx')

    # # data_sap = pd.read_excel(
    # #     'E:/Total/Station Data/Master Data/Data source/Data-SAP.xlsx')

    # writer_list = pd.ExcelWriter(path_list, engine='openpyxl')

    # df2.to_excel(
    #     writer_list, sheet_name='Station Data Brute', index=False)
    # writer_list.save()
    # writer_list.close()

    # print()

    data_sap = df1.copy()
    data_sharepoint = df2.copy()

    sh_p = data_sharepoint['Affiliate'].unique()
    sap_p = data_sap['Affiliate'].unique()

    for i in sh_p:

        if i in sap_p:

            element = i

            # print()

            # print('-'*20)
            # print(f"Pays : {element}")
            # print('-'*20)

            path_ecart = f"{folder_exp}/{element + '_' + str(today)}.xlsx"
            #path_list = f"{folder_list_affiliate}/list_affiliate_{str(today)}.xlsx"

            df_sap = data_sap[data_sap['Affiliate'] == element]
            df_sap.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
            df_sap = df_sap.drop_duplicates(subset="SAPCode", keep='first')
            dim_sap = df_sap.shape
            # print(f"dimension données SAP pour {element} est : {dim_sap}")
            df_sap['SAPCode'] = df_sap['SAPCode'].str.strip()

            df_sharepoint = data_sharepoint[data_sharepoint['Affiliate'] == element]
            df_sharepoint = df_sharepoint.drop_duplicates()
            dim_sharepoint = df_sharepoint.shape
            # print(f"dimension données sharepoint pour {element} est : {dim_sharepoint}")
            df_sharepoint['SAPCode'] = df_sharepoint['SAPCode'].str.strip()

            # print()

            # print("Comparaison :")
            # print('-'*7)

            X, Y, df_commun_1 = com(
                df_sap, df_sharepoint, col_name_1["text"], col_name_2["text"], texte=False)
            # a, cost, df_commun_2 = com(
            #     df_commun_1, df_sharepoint, 'SAPCode_BM', 'SAPCode_BM', texte=False)
            # b, cost, df_commun_3 = com(
            #     df_commun_2, df_sharepoint, 'SAPCode_BM_ISACTIVESITE', 'SAPCode_BM_ISACTIVESITE', texte=False)

            writer = pd.ExcelWriter(path_ecart, engine='openpyxl')
            df_sap.to_excel(writer, sheet_name='Data_SAP_Brute', index=False)
            df_sharepoint.to_excel(
                writer, sheet_name='Data_Sharepoint_Brute', index=False)
            X.to_excel(
                writer, sheet_name='ecart_SAP_vs_Sharepoint', index=False)
            Y.to_excel(
                writer, sheet_name='ecart_Sharepoint_vs_SAP', index=False)
            # a.to_excel(
            #     writer, sheet_name='SAP_vs_Sharepoint_SAPCode_BM', index=False)
            # b.to_excel(
            #     writer, sheet_name='SAP_vs_Sharepoint_SAPCode_BM_ISACTIVESITE', index=False)

            writer.save()
            writer.close()

            # print()
            # print('#'*70)
            # print()

            # sh = pd.read_excel("C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Data-SAP.xlsx")
            # sh = sh.drop_duplicates()
            # sh['SAPCode'] = sh['SAPCode'].str.strip()

            # z = sh['Affiliate'].unique()

            # # for w in z:
            # d = sh[sh['Affiliate']==w]

            # ecart_sap = X.copy()

            # ecart_sap = ecart_sap[["SAPCode", "Affiliate", "FINAL_SITENAME",
            #                        "SITETOWN", "ISACTIVESITE", "BUSINESSMODEL", "BM_source"]]
            # ecart_sap.columns = ['SAPCode', 'Affiliate', 'SAPName',
            #                      'Town', 'IsActiveSite', 'BUSINESSMODEL', 'BM_source']

            # colonnes = ['Zone', 'SubZone', 'IntermediateStatus', 'Brand', 'Segment', 'ContractMode', 'ShopSegment', 'SFSActivity', 'SFSContractType', 'PartnerOrBrand', 'TargetKit', 'TargetPOSprovider',
            #             'EstimatedInstallationDate', 'InstalledSolutionOnSite', 'SolutionProvider', 'SolutionInstallationDate', 'Status', 'SolutionRelease', 'SystemOwner', 'ConfigurationStatus',
            #             'IsAllPumpsConnectedToFCC', 'Reason', 'AutomaticTankGauging', 'ATGProvider', 'ATGModel', 'ATGConnected', 'ATGInstallationDate', 'TotalCardEPT connection', 'FuelCardProvider',
            #             'EPTHardware', 'EPTModel', 'EPTNumber', 'EPTConnected', 'PaymentLocation', 'HOSInstalled', 'HOSProvider', 'WSMSoftwareInstalled', 'WSMProvider', 'TELECOM', 'STABILITE TELECOM',
            #             'STARTBOXStatus',  'BM_source'
            #             ]

            # for col in colonnes:
            #     ecart_sap[col] = ""

            # all_cols_ordonner = ['SAPCode', 'Zone', 'SubZone', 'Affiliate', 'SAPName', 'Town',
            #                      'IsActiveSite', 'IntermediateStatus', 'Brand', 'Segment',
            #                      'BUSINESSMODEL', 'ContractMode', 'ShopSegment', 'SFSActivity',
            #                      'SFSContractType', 'PartnerOrBrand', 'TargetKit', 'TargetPOSprovider',
            #                      'EstimatedInstallationDate', 'InstalledSolutionOnSite',
            #                      'SolutionProvider', 'SolutionInstallationDate', 'Status',
            #                      'SolutionRelease', 'SystemOwner', 'ConfigurationStatus',
            #                      'IsAllPumpsConnectedToFCC', 'Reason', 'AutomaticTankGauging',
            #                      'ATGProvider', 'ATGModel', 'ATGConnected', 'ATGInstallationDate',
            #                      'TotalCardEPT connection', 'FuelCardProvider', 'EPTHardware',
            #                      'EPTModel', 'EPTNumber', 'EPTConnected', 'PaymentLocation',
            #                      'HOSInstalled', 'HOSProvider', 'WSMSoftwareInstalled', 'WSMProvider',
            #                      'TELECOM', 'STABILITE TELECOM', 'STARTBOXStatus', 'BM_source']

            # ecart_sap1 = ecart_sap.reindex(columns=all_cols_ordonner)
            # ecart_sap1['data_source'] = "ecart SAP"
            # ecart_sap1 = ecart_sap1[ecart_sap1['BUSINESSMODEL'] != 'CLOS']

            # sh = df_sharepoint.copy()

            # if a.shape[0] > 0:
            #     for j in range(a.shape[0]):
            #         for k in range(sh.shape[0]):
            #             if a['SAPCode'].iloc[j] == sh['SAPCode'].iloc[k]:
            #                 sh['BUSINESSMODEL'].iloc[k] = a['BUSINESSMODEL'].iloc[j]
            #                 sh['BM_source'].iloc[k] = a['BM_source'].iloc[j]

            # sh = sh[['SAPCode', 'Zone', 'SubZone', 'Affiliate', 'SAPName', 'Town', 'IsActiveSite', 'IntermediateStatus', 'Brand', 'Segment',
            #         'BUSINESSMODEL', 'ContractMode', 'ShopSegment', 'SFSActivity', 'SFSContractType', 'PartnerOrBrand', 'TargetKit', 'TargetPOSprovider',
            #          'EstimatedInstallationDate', 'InstalledSolutionOnSite', 'SolutionProvider', 'SolutionInstallationDate', 'Status',
            #          'SolutionRelease', 'SystemOwner', 'ConfigurationStatus', 'IsAllPumpsConnectedToFCC', 'Reason', 'AutomaticTankGauging',
            #          'ATGProvider', 'ATGModel', 'ATGConnected', 'ATGInstallationDate', 'TotalCardEPT connection', 'FuelCardProvider', 'EPTHardware',
            #          'EPTModel', 'EPTNumber', 'EPTConnected', 'PaymentLocation', 'HOSInstalled', 'HOSProvider', 'WSMSoftwareInstalled', 'WSMProvider',
            #          'TELECOM', 'STABILITE TELECOM', 'STARTBOXStatus', 'BM_source']]
            # sh['data_source'] = "Station Data"

            # sh_1 = sh.append(ecart_sap1, ignore_index=True)

            # book = load_workbook(path_list)
            # writer_list = pd.ExcelWriter(path_list, engine='openpyxl')
            # writer_list.book = book

            # sh_1.to_excel(writer_list, sheet_name=element, index=False)
            # writer_list.save()
            # writer_list.close()


# boucle principale
root.config(menu=mainMenu)
root.mainloop()
