import pandas as pd
import numpy as np
import os

from datetime import date, datetime
from openpyxl import load_workbook
from tqdm import tqdm


class EuroShare():
    """Cette class permet de comparer les données EuroDataHOS et Sharepoint et pour cela elle contient plusieures méthodes:

        Méthdoes:
        comperer : permet 
    """

    def __init__(self, path_data_HOS, path_data_sharepoint, path_Out):
        self.path_data_HOS = path_data_HOS
        self.path_data_sharepoint = path_data_sharepoint
        self.path_Out = path_Out

    def comparer(self, df_X, df_Y, col):

        self.diff_X = np.setdiff1d(df_X[col], df_Y[col])
        self.ecart_X = df_X.loc[df_X[col].isin(self.diff_X)]

        self.diff_Y = np.setdiff1d(df_Y[col], df_X[col])
        self.ecart_Y = df_Y.loc[df_Y[col].isin(self.diff_Y)]

        self.commun = df_X.loc[~df_X[col].isin(self.diff_X)]

        return self.ecart_X, self.ecart_Y, self.commun

    def LaodandPreprocess(self):

        # Chargement de Data EuroDataHOS et prétraitement
        self.df_hos = pd.read_excel(self.path_data_HOS)
        #df_hos.rename(columns={'SAPCODE': 'SAPCode'}, inplace=True)
        self.df_hos = self.df_hos.drop_duplicates(
            subset="SAPCode", keep='first')
        self.df_hos['SAPCode'] = self.df_hos['SAPCode'].str.strip()
        self.df_hos['SAPCode'] = self.df_hos['SAPCode'].astype(str)
        self.df_hos['FCC/POSsolution'] = self.df_hos['FCC / POSsolution'].str.strip()

        for h in range(self.df_hos.shape[0]):
            self.df_hos['FCC/POSsolution'].iloc[h] = self.df_hos['FCC/POSsolution'].iloc[h].split(" ")[
                0]
            if self.df_hos['Solution activée'].iloc[h] == "FCC + DMS-Shop":
                if self.df_hos['FCC/POSsolution'].iloc[h] == "FUELPOS":
                    self.df_hos['Corespo Installed Solution'].iloc[h] = "DMS-FCC-POS"
            if self.df_hos['Corespo Installed Solution'].iloc[h] in ["FCC-POS", "DMS-FCC-POS", "FCC-POS-BOS"]:
                self.df_hos['Corresp EPT connected'].iloc[h] = "Not Connected FCC-POS"

        # Chargement de Data Sharepoint et prétraitement
        self.df_sharepoint = pd.read_excel(self.path_data_sharepoint)
        self.df_sharepoint = self.df_sharepoint.drop_duplicates()
        self.df_sharepoint['SAPCode'] = self.df_sharepoint['SAPCode'].str.strip()
        self.df_sharepoint['SAPCode'] = self.df_sharepoint['SAPCode'].astype(
            str)
        self.df_sharepoint['EPTConnected'] = self.df_sharepoint['EPTConnected'].str.strip(
        )
        self.df_sharepoint['ATGConnected'] = self.df_sharepoint['ATGConnected'].str.strip(
        )
        self.df_sharepoint['ATGConnected'] = self.df_sharepoint['ATGConnected'].replace(
            "Not connected FCC", "Not Connected FCC")

        for s in range(self.df_sharepoint.shape[0]):
            if self.df_sharepoint['InstalledSolutionOnSite'].iloc[s] in ["FCC-POS", "DMS-FCC-POS", "FCC-POS-BOS", "DMS-FCC-POS-BOS"]:
                self.df_sharepoint['EPTConnected'].iloc[s] = "Not Connected FCC-POS"

        return self.df_hos, self.df_sharepoint

    def export_excel(self, path, df, SheetName):
        self.writer_list = pd.ExcelWriter(path, engine='openpyxl')
        df.to_excel(self.writer_list, sheet_name=SheetName, index=False)
        self.writer_list.save()
        self.writer_list.close()

    def export_excel_add_new_sheet(self, path, df, SheetName):
        self.book = load_workbook(path)
        self.writer_list = pd.ExcelWriter(path, engine='openpyxl')
        self.writer_list.book = self.book
        df.to_excel(self.writer_list, sheet_name=SheetName, index=False)
        self.writer_list.save()
        self.writer_list.close()

    def ecoder_InstalledSolutionOnSite(self, df):
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "DMS", "01- DMS#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "DMS-FCC", "01- DMS#02- FCC#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "DMS-FCC-POS", "01- DMS#02- FCC#03- POS#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "DMS-FCC-POS-BOS", "01- DMS#02- FCC#03- POS#04- BOS (Advanced/Premium)#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "FCC", "02- FCC#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "FCC-POS", "02- FCC#P03- POS#")
        df['InstalledSolutionOnSite'] = df['InstalledSolutionOnSite'].replace(
            "FCC-POS-BOS", "02- FCC#03- POS#04- BOS (Advanced/Premium)#")

    def UpdateSharepoint(self):

        for j in tqdm(range(self.df_commun_avec_sh.shape[0])):
            for k in range(self.df_hos.shape[0]):
                if self.df_commun_avec_sh['SAPCode'].iloc[j] == self.df_hos['SAPCode'].iloc[k]:
                    if self.df_commun_avec_sh['InstalledSolutionOnSite'].iloc[j] != self.df_hos['Corespo Installed Solution'].iloc[k]:
                        self.df_commun_avec_sh['InstalledSolutionOnSite'].iloc[
                            j] = self.df_hos['Corespo Installed Solution'].iloc[k]
                        self.df_commun_avec_sh['SolutionRelease'].iloc[j] = self.df_hos['FCC / POSsolution'].iloc[k]
                        self.df_commun_avec_sh['InstalledSolutionOnSite Source'].iloc[
                            j] = self.df_hos['InstalledSolutionOnSite Source'].iloc[k]

                if self.df_commun_avec_sh['SAPCode'].iloc[j] == self.df_hos['SAPCode'].iloc[k]:
                    if self.df_commun_avec_sh['EPTConnected'].iloc[j] != "":
                        if self.df_commun_avec_sh['EPTConnected'].iloc[j] != self.df_hos['Corresp EPT connected'].iloc[k]:
                            self.df_commun_avec_sh['EPTConnected'].iloc[j] = self.df_hos['Corresp EPT connected'].iloc[k]
                            self.df_commun_avec_sh['EPTConnected Source'].iloc[j] = self.df_hos['EPTConnected Source'].iloc[k]

                if self.df_commun_avec_sh['SAPCode'].iloc[j] == self.df_hos['SAPCode'].iloc[k]:
                    if self.df_commun_avec_sh['ATGConnected'].iloc[j] != "":
                        if self.df_commun_avec_sh['ATGConnected'].iloc[j] != self.df_hos['Coresp ATG Connected'].iloc[k]:
                            self.df_commun_avec_sh['ATGConnected'].iloc[j] = self.df_hos['Coresp ATG Connected'].iloc[k]
                            self.df_commun_avec_sh['ATGConnected Source'].iloc[j] = self.df_hos['ATGConnected Source'].iloc[k]

    def reduce(self):
        # charger et prétraiter les données EuroDataHOS et sharepoint
        self.df_hos, self.df_sharepoint = self.LaodandPreprocess()

        # Comparer les données sharepoint et EuroDataHOS
        self.X, self.Y, self.df_commun_avec_sh = self.comparer(
            self.df_sharepoint, self.df_hos, 'SAPCode')

        # Transformer la colonnes InstalledSolutionOnSite de sharepoint (ex: "DMS" => "01- DMS#")
        self.ecoder_InstalledSolutionOnSite(self.df_sharepoint)

        # Exporter les données EuroDataHOS
        self.export_excel(self.path_Out, self.df_hos, "EuroDataHOS")

        # Exporter les données sharepoint
        self.export_excel_add_new_sheet(
            self.path_Out, self.df_sharepoint, "Sharepoint")

        os.system('cls' if os.name == 'nt' else 'clear')
        print()
        print("-"*23)
        print("Traitement en cours...")
        print("-"*23)

        # Mettre à jour le données sharepoint avec le données EuroDataHOS
        self.UpdateSharepoint()

        # Transformer la colonnes InstalledSolutionOnSite de la table MAJ (ex: "DMS" => "01- DMS#")
        self.ecoder_InstalledSolutionOnSite(self.df_commun_avec_sh)

        # Exporter les données MAJ
        self.export_excel_add_new_sheet(
            self.path_Out, self.df_commun_avec_sh, "maj_sharepoint")


Week = 48.1
today = date.today().strftime("%d%m%y")
path_data_HOS = "E:/AppTotalEnergies/SRC-Python-comparaison/InputData/DataAppWeek48.xlsx"
path_data_sharepoint = "E:/AppTotalEnergies/SRC-Python-comparaison/InputData/sharepoint.xlsx"
path_Out = f"E:/AppTotalEnergies/SRC-Python-comparaison/OutputData/KPI-SIS-AFRIQUE-S{Week}-{today}.xlsx"

start = datetime.now()
m = EuroShare(path_data_HOS, path_data_sharepoint, path_Out)
m.reduce()

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
