from eurodatahos_vs_shrepoint import EuroShare
from datetime import date, datetime
import os


Week = 48.1
today = date.today().strftime("%d%m%y")
path_data_HOS = "E:/AppTotalEnergies/SRC-Python-comparaison/InputData/DataAppWeek48.xlsx"
path_data_sharepoint = "E:/AppTotalEnergies/SRC-Python-comparaison/InputData/sharepoint.xlsx"
path_Out = f"E:/AppTotalEnergies/SRC-Python-comparaison/OutputData/KPI-SIS-AFRIQUE-S{Week}-{today}.xlsx"

start = datetime.now()
m = EuroShare(path_data_HOS, path_data_sharepoint, path_Out)
m.reduce()

# os.system('cls' if os.name == 'nt' else 'clear')
print()
print("--------------------")
print("Terminer avec succès")
print("--------------------")
print()

end = datetime.now()
tm = end - start
print("temps d'exécution :", tm)
print()