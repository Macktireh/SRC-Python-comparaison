# importer la binliothèque
import tkinter as tk
from typing import Sized
import pandas as pd
import numpy as np
import pyttsx3
import os
import shutil
import time

from tkinter import Widget, filedialog, messagebox, ttk
from tkinter.constants import ACTIVE, END, TRUE
from datetime import date, datetime
from openpyxl import load_workbook


###########################################################################################
# --------------------------------------  Backend  -------------------------------------- #
###########################################################################################
#"C:\Users\J1049122\Desktop\Station Data\Master-Data\Data source"

def File_dialog_1():
    """Cette fonction ouvrira l'explorateur de fichiers et affectera le chemin de fichier choisi à label_file"""
    filename_1 = filedialog.askopenfilename(initialdir="C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source",
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

    frame1 = tk.LabelFrame(new_interface, text=label_file_1["text"])
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
    filename_2 = filedialog.askopenfilename(initialdir="C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source",
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

    frame2 = tk.LabelFrame(new_interface, text=label_file_2["text"])
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
        print(f"il y'a {len(diff_X)} de différence")

        print()
        diff_Y = np.setdiff1d(df_Y[col_y], df_X[col_x])
        ecart_Y = df_Y.loc[df_Y[col_y].isin(diff_Y)]

        print("Données Sharepoint versus données SAP :")
        print(f"il y'a {len(diff_Y)} de différence")

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

    # today = datetime.now().strftime("%d-%m-%Y_%Hh-%Mmin-%Ss")
    today = datetime.now().strftime("%d%m%y")
    if VarNameFloder.get() == "" or None:
        folder_result = "{}/resultats_{}".format(path_folder_result["text"], today)
    else:
        folder_result = "{}/{}_{}".format(path_folder_result["text"],VarNameFloder.get(), today)

    if os.path.exists(folder_result):
        shutil.rmtree(f'{folder_result}')
        print(
            f"le dossier {folder_result} à été bien supprimer et recréer\n-------------")
        print()
    else:
        print(f"le dossier {folder_result} n'existe pas\n-------------")
        print()

    os.mkdir(folder_result)

    def Export_ExcelSheet(path, df_1, df_2, df_3, df_4, sn_1="fichier_1", sn_2="fichier_2"):

        writer = pd.ExcelWriter(path, engine='openpyxl')
        df_1.to_excel(writer, sheet_name=sn_1, index=False)
        df_2.to_excel(
        writer, sheet_name=sn_2, index=False)

        df_3.to_excel(writer, sheet_name=f'ecart {sn_1} vs {sn_2}', index=False)
        df_4.to_excel(writer, sheet_name=f'ecart {sn_2} vs {sn_1}', index=False)
        writer.save()
        writer.close()
        

    # folder_exp = f'{folder_result}/AFR_{today}'

    # if os.path.exists(folder_exp):
    #     shutil.rmtree(f'{folder_exp}')
    #     print(
    #         f"le dossier AFR_{today} à été bien supprimer et recréer\n-------------")
    #     print()
    # else:
    #     print(f"le dossier AFR_{today} n'existe pas\n-------------")
    #     print()

    # os.mkdir(folder_exp)

    dataframe_1 = df1.copy()
    dataframe_2 = df2.copy()

    liste_pays_1 = dataframe_1['Affiliate'].unique()
    liste_pays_2 = dataframe_2['Affiliate'].unique()

    if VarCheckButton_for_country.get() == True:

        for i in liste_pays_2:

            if i in liste_pays_1:

                element = i

                print()

                print('-'*20)
                print(f"Pays : {element}")
                print('-'*20)

                # path_ecart = f"{folder_result}/{element + '_' + str(today)}.xlsx"
                path_ecart = f"{folder_result}/{element}.xlsx"
                #path_list = f"{folder_list_affiliate}/list_affiliate_{str(today)}.xlsx"

                df_1 = dataframe_1[dataframe_1['Affiliate'] == element]
                #df_sap.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
                if VarCheckButton_Duplicates.get() == TRUE:
                    df_1 = df_1.drop_duplicates(subset=col_name_1["text"], keep='first')
                dim_df_1 = df_1.shape
                print(f"dimension données SAP pour {element} est : {dim_df_1}")
                df_1[col_name_1["text"]] = df_1[col_name_1["text"]].str.strip()

                df_2 = dataframe_2[dataframe_2['Affiliate'] == element]
                if VarCheckButton_Duplicates.get() == TRUE:
                    df_2 = df_2.drop_duplicates(subset=col_name_2["text"], keep='first')
                dim_df_2 = df_2.shape
                print(
                    f"dimension données sharepoint pour {element} est : {dim_df_2}")
                df_2[col_name_2["text"]] = df_2[col_name_2["text"]].str.strip()

                print()

                print("Comparaison :")
                print('-'*7)

                ecat_1, ecat_2, df_commun_1 = com(
                    df_1, df_2, col_name_1["text"], col_name_2["text"])

                # writer = pd.ExcelWriter(path_ecart, engine='openpyxl')
                # df_1.to_excel(writer, sheet_name='fichier_1', index=False)
                # df_2.to_excel(
                #     writer, sheet_name='fichier_2', index=False)
                # ecat_1.to_excel(
                #     writer, sheet_name='ecart fichier_1 vs fichier_2', index=False)
                # ecat_2.to_excel(
                #     writer, sheet_name='ecart fichier_1 vs fichier_2', index=False)

                # writer.save()
                # writer.close()
                if (str(VarNameExportSheet_1.get()) == "") & (str(VarNameExportSheet_2.get()) == "") :
                    Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2)

                elif (str(VarNameExportSheet_1.get()) != "") & (str(VarNameExportSheet_2.get()) != "") :
                    Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                        sn_1=str(VarNameExportSheet_1.get()),
                                        sn_2=str(VarNameExportSheet_2.get()))

                elif str(VarNameExportSheet_1.get()) != "" :
                    Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                        sn_1=str(VarNameExportSheet_1.get()))

                elif str(VarNameExportSheet_2.get()) != "" :
                    Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                        sn_2=str(VarNameExportSheet_2.get()))

        msg_info = messagebox.showinfo("Information", "Succès")
    else:
        element = VarNameFloder.get()
        path_ecart = f"{folder_result}/{element}.xlsx"

        df_1 = dataframe_1
        if VarCheckButton_Duplicates.get() == TRUE:
            df_1 = df_1.drop_duplicates(subset=col_name_1["text"], keep='first')
        dim_df_1 = df_1.shape
        print(f"dimension données SAP pour {element} est : {dim_df_1}")
        df_1[col_name_1["text"]] = df_1[col_name_1["text"]].str.strip()

        df_2 = dataframe_2
        if VarCheckButton_Duplicates.get() == TRUE:
            df_2 = df_2.drop_duplicates(subset=col_name_2["text"], keep='first')
        dim_df_2 = df_2.shape
        print(
            f"dimension données sharepoint pour {element} est : {dim_df_2}")
        df_2[col_name_2["text"]] = df_2[col_name_2["text"]].str.strip()

        print()

        print("Comparaison :")
        print('-'*7)

        ecat_1, ecat_2, df_commun_1 = com(df_1, df_2, col_name_1["text"], col_name_2["text"])

        if (str(VarNameExportSheet_1.get()) == "") & (str(VarNameExportSheet_2.get()) == "") :
            Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2)

        elif (str(VarNameExportSheet_1.get()) != "") & (str(VarNameExportSheet_2.get()) != "") :
            Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                sn_1=str(VarNameExportSheet_1.get()),
                                sn_2=str(VarNameExportSheet_2.get()))

        elif str(VarNameExportSheet_1.get()) != "" :
            Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                sn_1=str(VarNameExportSheet_1.get()))

        elif str(VarNameExportSheet_2.get()) != "" :
            Export_ExcelSheet(path_ecart, df_1, df_2, ecat_1, ecat_2,
                                sn_2=str(VarNameExportSheet_2.get()))
        
        msg_info = messagebox.showinfo("Information", "Succès")
                


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
    filename = filedialog.askdirectory(initialdir="C:/Users/J1049122/Desktop/Station Data/Master-Data")
    path_folder_result["text"] = filename


def CheckButton1():
    if VarCheckButton_1.get():
        sheet_name_1['state'] = "normal"
    else:
        sheet_name_1['state'] = "disabled"
        sheet_name_1['disabledbackground'] = '#9F9D99'


def CheckButton2():
    if VarCheckButton_2.get():
        sheet_name_2['state'] = "normal"
    else:
        sheet_name_2['state'] = "disabled"
        sheet_name_2['disabledbackground'] = '#9F9D99'

def ClearAll():
    # pass
    col_name_1["text"] = ""
    col_name_2["text"] = ""

    label_file_1['text'] = "Aucun fichier sélectionné"
    label_file_2['text'] = "Aucun fichier sélectionné"

    box1.delete(0, END)
    box2.delete(0, END)

    VarCheckButton_1.set(False)
    VarCheckButton_2.set(False)

    sheet_name_1.delete(0, END)
    sheet_name_1['state'] = "disabled"
    sheet_name_2.delete(0, END)
    sheet_name_2['state'] = "disabled"
    
    path_folder_result['text'] = ""

    VarNameFloder.set("")




############################################################################################
# --------------------------------------  Frontend  -------------------------------------- #
############################################################################################


# création de l'objet de la fenetre
root = tk.Tk()

# personnaliser la fenetre
root.title("    PyApp Station Data Desktop")  # nom d'entête de la fenetre
root.iconbitmap("TotalEnergies.ico")  # icone de la fenetre
root.geometry("1080x600+15+15")  # taille de la fenetre
root.minsize(1040, 620)
root.maxsize(1120, 800)

# configuration du font de la fenetre (couleur ou autre)
# root.config(background='#CCCCCC')

# barre de menu
mainMenu = tk.Menu(root)

file_menu = tk.Menu(root, tearoff=0)
file_menu.add_command(label="Parcourir le fichier 1", command=File_dialog_1)
file_menu.add_command(label="Parcourir le fichier 2", command=File_dialog_2)
file_menu.add_command(label="Quit", command=root.quit)

mainMenu.add_cascade(label="File", menu=file_menu)


# ---------- la boîte de dialogue d'ouverture de fichier ---------- #
file_frame_1 = tk.LabelFrame(
    root, text="Ouvrir le premier fichier", background='#CCCCCC')
file_frame_1.place(height=200, width=500, rely=0.05, relx=0.01)

# label
fram_label_1 = tk.Frame(file_frame_1, background='#CCCCCC')

VarCheckButton_1 = tk.BooleanVar()
VarCheckButton_1.set(False)
CheckButton_1 = tk.Checkbutton(
    fram_label_1, var=VarCheckButton_1, command=CheckButton1, background='#CCCCCC')
CheckButton_1.grid(row=0, column=0)

label_1 = tk.Label(fram_label_1, background='#CCCCCC',
                   text="Cochez et indiquez le nom de la feuille d'Excel par défaut premier feuille (facultatif)")
label_1.grid(row=0, column=1)

var_entry_1 = tk.StringVar()
sheet_name_1 = tk.Entry(file_frame_1, textvariable=var_entry_1, bd=2,
                        background="white", state="disabled", disabledbackground='#9F9D99')
sheet_name_1['state'] = "disabled"
sheet_name_1.place(rely=0.63, relx=0.10)

fram_label_1.place(rely=0.45, relx=0)

# Buttons
button1 = tk.Button(file_frame_1, text="Parcourir",
                    command=lambda: File_dialog_1())
button1.place(rely=0.83, relx=0.50)

button2 = tk.Button(file_frame_1, text="Charger",
                    command=lambda: view_data())
button2.place(rely=0.83, relx=0.30)


# Le texte du fichier/chemin d'accès au fichier
label_file_1 = ttk.Label(file_frame_1, text="Aucun fichier sélectionné")
label_file_1.place(rely=0.02, relx=0)

path_import_frame_1 = tk.Frame(file_frame_1, background='#CCCCCC')

lbl_name_export_sheet_1 = tk.Label(path_import_frame_1,text="Renomer le fichier (facultatif) ", background='#CCCCCC')
lbl_name_export_sheet_1.grid(row=0, column=0)

VarNameExportSheet_1 = tk.StringVar()
entry_name_export_sheet_1 = tk.Entry(path_import_frame_1, textvariable=VarNameExportSheet_1, width=25)
entry_name_export_sheet_1.grid(row=0, column=1)

path_import_frame_1.place(rely=0.15, relx=0)


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
                      background="#F1F2F6", width=30)
col_name_1.place(rely=0.53, relx=0.24)

btn_1 = tk.Button(root, text='Ok', command=selected_item_1)
btn_1.place(rely=0.63, relx=0.28)


######## ---------- la boîte de dialogue d'ouverture de fichier ---------- #########


file_frame_2 = tk.LabelFrame(
    root, text="Ouvrir le deuxième fichier", background='#CCCCCC')
file_frame_2.place(height=200, width=500, rely=0.05, relx=0.51)

# label
fram_label_2 = tk.Frame(file_frame_2, background='#CCCCCC')

VarCheckButton_2 = tk.BooleanVar()
VarCheckButton_2.set(False)
CheckButton_2 = tk.Checkbutton(
    fram_label_2, var=VarCheckButton_2, command=CheckButton2, background='#CCCCCC')
CheckButton_2.grid(row=0, column=0)

label_2 = tk.Label(fram_label_2, background='#CCCCCC',
                   text="Cochez et indiquez le nom de la feuille d'Excel par défaut premier feuille (facultatif)")
label_2.grid(row=0, column=1)

var_entry_2 = tk.StringVar()
sheet_name_2 = tk.Entry(file_frame_2, textvariable=var_entry_2, bd=2,
                        background="white", state="disabled", disabledbackground='#9F9D99')
sheet_name_2.place(rely=0.63, relx=0.10)

fram_label_2.place(rely=0.45, relx=0)

# Buttons
button3 = tk.Button(file_frame_2, text="Parcourir",
                    command=lambda: File_dialog_2())
button3.place(rely=0.83, relx=0.50)

button4 = tk.Button(file_frame_2, text="Charger",
                    command=lambda: view_data_2())
button4.place(rely=0.83, relx=0.30)

# Le texte du fichier/chemin d'accès au fichier
label_file_2 = ttk.Label(file_frame_2, text="Aucun fichier sélectionné")
label_file_2.place(rely=0.02, relx=0)

path_import_frame_2 = tk.Frame(file_frame_2, background='#CCCCCC')

lbl_name_export_sheet_2 = tk.Label(path_import_frame_2,text="Renomer le fichier (facultatif) ", background='#CCCCCC')
lbl_name_export_sheet_2.grid(row=0, column=0)

VarNameExportSheet_2 = tk.StringVar()
entry_name_export_sheet_2 = tk.Entry(path_import_frame_2, textvariable=VarNameExportSheet_2, width=25)
entry_name_export_sheet_2.grid(row=0, column=1)

path_import_frame_2.place(rely=0.15, relx=0)


# Liste des colonnes
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
                      background="#F1F2F6", width=30)
col_name_2.place(rely=0.53, relx=0.72)

btn_2 = tk.Button(root, text='Ok', command=selected_item_2)
btn_2.place(rely=0.63, relx=0.76)


# ---------------------------------------------------------------- #


button_comparer = tk.Button(
    root, text="Comparer", width=20, background="#004C8C", fg="white", command=compare)
button_comparer.place(rely=0.90, relx=0.25)

button_reunisialise = tk.Button(
    root, text="Réinitialiser", width=20, background="#FF9933", fg="white", command=ClearAll)
button_reunisialise.place(rely=0.90, relx=0.45)

button_quit = tk.Button(
    root, text="Quiter", width=20, background="#C60030", fg="white", command=root.quit)
button_quit.place(rely=0.90, relx=0.65)


fram = tk.Frame(root, bd=1)
# folder_path = tk.StringVar()
# lbl1 = tk.Label(fram, textvariable=folder_path)
VarCheckButton_Duplicates = tk.BooleanVar()
VarCheckButton_Duplicates.set(TRUE)
CheckButton_Duplicates = tk.Checkbutton(fram, var=VarCheckButton_Duplicates)
CheckButton_Duplicates.grid(row=0, column=0, padx=0, ipadx=0)

lbl_Duplicates = tk.Label(fram, text="Cohez pour enlèver les doublons avant de comparer")
lbl_Duplicates.grid(row=0, column=1, padx=0, ipadx=0)

button_folder = tk.Button(
    fram, text="Sélectionner un dossier de destination", command=browse_button)
button_folder.grid(row=0, column=2)

path_folder_result = tk.Label(fram, text="")
path_folder_result.grid(row=0, column=3)

VarCheckButton_for_country = tk.BooleanVar()
VarCheckButton_for_country.set(TRUE)
CheckButton_for_country = tk.Checkbutton(fram, var=VarCheckButton_for_country)
CheckButton_for_country.grid(row=1, column=0)

lbl_name_folder = tk.Label(fram, text="Nom du dossier de reslutat :")
lbl_name_folder.grid(row=1, column=1)

VarNameFloder = tk.StringVar()
name_folder_result = tk.Entry(fram, textvariable=VarNameFloder)
name_folder_result.grid(row=1, column=2)

fram.place(rely=0.80, relx=0.02)


# button_comparer = tk.Button(
#     root, text="compare", width=20, background="#3FB8F2", command=lambda: compare)
# button_folder.place(rely=0.90, relx=0.40)

# boucle principale
root.config(menu=mainMenu)
root.mainloop()