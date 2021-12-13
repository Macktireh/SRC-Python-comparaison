if True:
    import tkinter as tk
    import pandas as pd
    import numpy as np
    from tkinter import filedialog, messagebox, ttk, PhotoImage
    from datetime import date, datetime
    from PIL import Image, ImageTk


class AppUI:

    def __init__(self):
        root = tk.Tk()
        self.root = root
        self.root.title("Data App Desktop")
        self.root.iconbitmap("media/TotalEnergies.ico")
        self.root.config(background="#BCBCBC")
        
        width, height = 550, 300
        frm_width = self.root.winfo_rootx() - self.root.winfo_x()
        win_width = width + 2 * frm_width
        titlebar_height = self.root.winfo_rooty() - self.root.winfo_y()
        win_height = (height+10) + titlebar_height + frm_width
        x = self.root.winfo_screenwidth() // 2 - win_width // 2
        y = self.root.winfo_screenheight() // 2 - win_height // 2
        self.root.geometry("{}x{}+{}+{}".format(width, height, x, y))
        
        # charger les images
        self.excelIcon = PhotoImage(file="media/excel.png")
        self.excelIcon = self.excelIcon.subsample(10, 10)
        self.csvIcon = PhotoImage(file="media/csv.png")
        self.csvIcon = self.csvIcon.subsample(10, 10)
        self.viewIcon = PhotoImage(file="media/view.png")
        self.viewIcon = self.viewIcon.subsample(50, 50)

        super().__init__()
        self.Home()

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
            command=self.Window_SAP_vs_Sharepoint)
        self.Btn_SAP_vs_Sharepoint.place(relx=0.09, rely=0.2, relheight=0.15, relwidth=0.4)

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
            command=self.Window_EuroDataHos_vs_Sharepoint)
        self.Btn_EuroDataHOS_vs_Sharepoint.place(relx=0.51, rely=0.2, relheight=0.15, relwidth=0.4)

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
            command=self.root.quit)
        self.BtnExit.place(relx=0.35, rely=0.6, relheight=0.15, relwidth=0.3)
        
    def Container(self, win, title, position=0):
        
        self.Frame = tk.LabelFrame(
            win, 
            text=title, 
            font=("Helvetica 10 bold"), 
            fg="#004C8C", labelanchor='n')
        if position == 0:
            self.Frame.place(relx=0.02, rely=0.08, relheight=0.35, relwidth=0.96)
        else:
            self.Frame.place(relx=0.02, rely=0.45, relheight=0.35, relwidth=0.96)
            
        self.VarLabelPath = tk.StringVar()
        self.VarLabelPath.set("bla bla bla")
        self.LabelPath = tk.Label(self.Frame, textvariable=self.VarLabelPath, bg="#FAEBD7")
        self.LabelPath.pack(fill="x")
        
        # Button import avec icon
        self.excelBtn = tk.Button(
            self.Frame,
            image=self.excelIcon,
            text="Import data from Excel",
            compound="top",
            height=70,
            width=160,
            bd=1,
            bg="#DCDCDC",
            command=None,
            pady=2
        ).place(relx=0.23, rely=0.21)

        self.csvBtn = tk.Button(
            self.Frame,
            image=self.csvIcon,
            text="Import data from CSV",
            compound="top",
            height=70,
            width=160,
            bd=1,
            bg="#DCDCDC",
            command=None,
        ).place(relx=0.53, rely=0.21)
    
        self.ViewDataBtn = tk.Button(
            self.Frame,
            image=self.viewIcon,
            text=f"   Voir le données {title}",
            compound="left",
            height=20,
            width=190,
            bd=1,
            bg="#DCDCDC",
            command=None,
        ).place(relx=0.35, rely=0.8)
    
    def ContainerBtn(self, win, funct=None):
        self.FrameBtn = tk.LabelFrame(win)
        self.FrameBtn.place(relx=0.02, rely=0.82, relheight=0.15, relwidth=0.96)

        # Button sortie
        self.BtnSortie = tk.Button(
            self.FrameBtn, 
            text="Sortie", 
            font=("Helvetica", 11),
            width=15,
            height=1,
            borderwidth=5,
            relief="raised",
            command=None)
        self.BtnSortie.place(relx=0.1, rely=0.2)
        
        # Button Comparer
        self.BtnComparer = tk.Button(
            self.FrameBtn, 
            text="Lancer", 
            font=("Helvetica", 11),
            width=15,
            height=1,
            borderwidth=5,
            relief="raised",
            background="#004C8C",
            activebackground="#004C8C",
            foreground="black",
            activeforeground="white",
            command=funct)
        self.BtnComparer.place(relx=0.38, rely=0.2)
        
        # Button Fermer
        self.BtnFermer = tk.Button(
            self.FrameBtn, 
            text="Fermer", 
            font=("Helvetica", 11),
            width=15,
            height=1,
            borderwidth=5,
            relief="raised",
            background="#C60030",
            activebackground="#C60030",
            foreground="black",
            activeforeground="white",
            command=self.TopWindow.destroy)
        self.BtnFermer.place(relx=0.65, rely=0.2)
    
    def Toplevel_Window(self, title, source):
        self.TopWindow = tk.Toplevel(self.root)
        self.TopWindow.grab_set()
        self.TopWindow.title("Data App Desktop")
        self.TopWindow.iconbitmap("media/TotalEnergies.ico")
        self.TopWindow.geometry("700x500+15+15")
        self.TopWindow.resizable(width=False, height=False)
        
        self.header = tk.Frame(self.TopWindow, bd=4, bg="#FAEBD7", height=5)
        self.header.pack(side="top", fill="x")
        
        self.title = tk.Label(
            self.header,
            text=title,
            font=("Helvetica", 15),
            bg="#FAEBD7",)
        self.title.pack(side="bottom", fill="x")
        
        self.Container(self.TopWindow, source)
        self.Container(self.TopWindow, 'Sharepoint', position=1)
        self.ContainerBtn(self.TopWindow)

    def Window_SAP_vs_Sharepoint(self):
        self.Toplevel_Window("SAP versus Sharepoint", "SAP")
        
    def Window_EuroDataHos_vs_Sharepoint(self):
        self.Toplevel_Window("EuroDataHos versus Sharepoint", "EuroDataHOS")