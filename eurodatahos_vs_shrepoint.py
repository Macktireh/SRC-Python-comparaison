import pandas as pd
import numpy as np
import pyttsx3
import os
import shutil
import time

from datetime import date, datetime
from openpyxl import load_workbook
from tqdm import tqdm

start = datetime.now()
Week = 48

path_data_HOS = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Comparaison Appro & Sharepoint/DataAppWeek48.xlsx"
path_data_sharepoint = "C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Comparaison Appro & Sharepoint/sharepoint.xlsx"
path_Out = f"C:/Users/J1049122/Desktop/Station Data/Master-Data/Data source/Comparaison Appro & Sharepoint/KPI-SIS-AFRIQUE-S{Week}.xlsx"

def com(df_X, df_Y, col):

    diff_X = np.setdiff1d(df_X[col] ,df_Y[col])
    ecart_X = df_X.loc[df_X[col].isin(diff_X)]

    diff_Y = np.setdiff1d(df_Y[col], df_X[col])
    ecart_Y = df_Y.loc[df_Y[col].isin(diff_Y)]
    
    commun = df_X.loc[~df_X[col].isin(diff_X)]

    return ecart_X, ecart_Y, commun     

def export_excel(path, df, SheetName):
    writer_list = pd.ExcelWriter(path, engine = 'openpyxl')
    df.to_excel(writer_list, sheet_name = SheetName, index=False)
    writer_list.save()
    writer_list.close()

def export_excel_add_new_sheet(path, df, SheetName):
    book = load_workbook(path)
    writer_list = pd.ExcelWriter(path, engine = 'openpyxl')
    writer_list.book = book
    df.to_excel(writer_list, sheet_name = SheetName, index=False)
    writer_list.save()
    writer_list.close()

def ecoder_InstalledSolutionOnSite(df):
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("DMS", "01- DMS#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("DMS-FCC", "01- DMS#02- FCC#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("DMS-FCC-POS", "01- DMS#02- FCC#03- POS#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("DMS-FCC-POS-BOS", "01- DMS#02- FCC#03- POS#04- BOS (Advanced/Premium)#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("FCC", "02- FCC#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("FCC-POS", "02- FCC#P03- POS#")
    df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace("FCC-POS-BOS", "02- FCC#03- POS#04- BOS (Advanced/Premium)#")

def comparer():

    df_hos = pd.read_excel(path_data_HOS)
    # df_hos.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
    df_hos = df_hos.drop_duplicates(subset = "SAPCode", keep = 'first')
    df_hos['SAPCode'] = df_hos['SAPCode'].str.strip()
    df_hos['SAPCode'] = df_hos['SAPCode'].astype(str)
    df_hos['FCC/POSsolution'] = df_hos['FCC / POSsolution'].str.strip()

    for h in range(df_hos.shape[0]):
        df_hos['FCC/POSsolution'].iloc[h] = df_hos['FCC/POSsolution'].iloc[h].split(" ")[0]
        if df_hos['Solution activée'].iloc[h] == "FCC + DMS-Shop":
            if df_hos['FCC/POSsolution'].iloc[h] == "FUELPOS":
                df_hos['Corespo Installed Solution'].iloc[h] = "DMS-FCC-POS"
        if df_hos['Corespo Installed Solution'].iloc[h] in ["FCC-POS", "DMS-FCC-POS", "FCC-POS-BOS"]:
            df_hos['Corresp EPT connected'].iloc[h] = "Not Connected FCC-POS"


    df_sharepoint = pd.read_excel(path_data_sharepoint)
    df_sharepoint = df_sharepoint.drop_duplicates()
    df_sharepoint['SAPCode'] = df_sharepoint['SAPCode'].str.strip()
    df_sharepoint['SAPCode'] = df_sharepoint['SAPCode'].astype(str)
    df_sharepoint['EPTConnected'] = df_sharepoint['EPTConnected'].str.strip()
    df_sharepoint['ATGConnected'] = df_sharepoint['ATGConnected'].str.strip()
    df_sharepoint['ATGConnected'] = df_sharepoint['ATGConnected'].replace("Not connected FCC", "Not Connected FCC")


    for s in range(df_sharepoint.shape[0]):
        if df_sharepoint['InstalledSolutionOnSite'].iloc[s] in ["FCC-POS", "DMS-FCC-POS", "FCC-POS-BOS", "DMS-FCC-POS-BOS"]:
            df_sharepoint['EPTConnected'].iloc[s] = "Not Connected FCC-POS"

    X, Y, df_commun_avec_sh = com(df_sharepoint, df_hos, 'SAPCode')

    ecoder_InstalledSolutionOnSite(df_sharepoint)
    export_excel(path_Out, df_hos, "EuroDataHOS")
    export_excel_add_new_sheet(path_Out, df_sharepoint, "Sharepoint")

    os.system('cls' if os.name == 'nt' else 'clear')
    print()
    print("-"*23)
    print("Traitement en cours...")
    print("-"*23)

    for j in tqdm(range(df_commun_avec_sh.shape[0])):
        for k in range(df_hos.shape[0]):
            if df_commun_avec_sh['SAPCode'].iloc[j] == df_hos['SAPCode'].iloc[k]:
                if df_commun_avec_sh['InstalledSolutionOnSite'].iloc[j] != df_hos['Corespo Installed Solution'].iloc[k]:
                    df_commun_avec_sh['InstalledSolutionOnSite'].iloc[j] = df_hos['Corespo Installed Solution'].iloc[k]
                    df_commun_avec_sh['SolutionRelease'].iloc[j] = df_hos['FCC / POSsolution'].iloc[k]
                    df_commun_avec_sh['InstalledSolutionOnSite Source'].iloc[j] = df_hos['InstalledSolutionOnSite Source'].iloc[k]

            if df_commun_avec_sh['SAPCode'].iloc[j] == df_hos['SAPCode'].iloc[k]:
                if df_commun_avec_sh['EPTConnected'].iloc[j] != "":
                    if df_commun_avec_sh['EPTConnected'].iloc[j] != df_hos['Corresp EPT connected'].iloc[k]:
                        df_commun_avec_sh['EPTConnected'].iloc[j] = df_hos['Corresp EPT connected'].iloc[k]
                        df_commun_avec_sh['EPTConnected Source'].iloc[j] = df_hos['EPTConnected Source'].iloc[k]
                    
            if df_commun_avec_sh['SAPCode'].iloc[j] == df_hos['SAPCode'].iloc[k]:
                if df_commun_avec_sh['ATGConnected'].iloc[j] != "":
                    if df_commun_avec_sh['ATGConnected'].iloc[j] != df_hos['Coresp ATG Connected'].iloc[k]:
                        df_commun_avec_sh['ATGConnected'].iloc[j] = df_hos['Coresp ATG Connected'].iloc[k]
                        df_commun_avec_sh['ATGConnected Source'].iloc[j] = df_hos['ATGConnected Source'].iloc[k]

    ecoder_InstalledSolutionOnSite(df_commun_avec_sh)
    export_excel_add_new_sheet(path_Out, df_commun_avec_sh, "maj_sharepoint")


comparer()

os.system('cls' if os.name == 'nt' else 'clear')
print()
print("--------------------")
print("Terminer avec succès")
print("--------------------")
print()

end = datetime.now()
tm = end - start
print("temps d'exécution :", tm)
print()