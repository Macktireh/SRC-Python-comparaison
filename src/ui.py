if True:
    import tkinter as tk
    import pandas as pd
    import numpy as np
    import os
    import shutil
    from tkinter import filedialog, messagebox, ttk, PhotoImage
    from tkinter.constants import ACTIVE, END, RAISED, TRUE
    from datetime import date, datetime
    from openpyxl import load_workbook
    from PIL import Image, ImageTk


class AppUI:

    def __init__(self):
        root = tk.Tk()
        self.root = root
        # self.root.withdraw()
        self.root.title("Data App Desktop")
        self.root.geometry("550x300")
        self.root.iconbitmap("media/TotalEnergies.ico")
        self.root.config(background="#BCBCBC")
        self.root.maxsize(600, 330)
        self.root.minsize(500, 280)

        super().__init__()
        self.Home()

    def onExit(self):
        self.root.quit()

    def Home(self):

        self.Btn_SAP_vs_Sharepoint = tk.Button(
            self.root,
            text="SAP vs Sharepoint",
            background="#FAEBD7",
            activebackground="#0256CD",
            foreground="black",
            activeforeground="white",
            borderwidth=5,
            relief="raised",
            font=("Helvetica", 10),
            command=None)
        self.Btn_SAP_vs_Sharepoint.place(
            relx=0.09, rely=0.2, relheight=0.15, relwidth=0.4)

        self.Btn_EuroDataHOS_vs_Sharepoint = tk.Button(
            self.root,
            text="EuroDataHOS vs Sharepoint",
            background="#FAEBD7",
            activebackground="#0256CD",
            foreground="black",
            activeforeground="white",
            borderwidth=5,
            relief="raised",
            font=("Helvetica", 10),
            command=None)
        self.Btn_EuroDataHOS_vs_Sharepoint.place(
            relx=0.51, rely=0.2, relheight=0.15, relwidth=0.4)

        self.BtnExit = tk.Button(
            self.root,
            text="Quiter",
            background="#C60030",
            activebackground="#C60030",
            foreground="black",
            activeforeground="white",
            borderwidth=5,
            relief="raised",
            font=("Helvetica", 11),
            command=self.onExit)
        self.BtnExit.place(
            relx=0.35, rely=0.6, relheight=0.15, relwidth=0.3)
