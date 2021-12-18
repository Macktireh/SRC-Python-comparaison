if True:
    import tkinter as tk
    import pandas as pd
    import numpy as np
    from tkinter import filedialog, messagebox, ttk, PhotoImage
    from datetime import date, datetime
    from PIL import Image, ImageTk
    from eurodatahos_vs_shrepoint import EuroShare


class AppUI():

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
        
        self.typefile = None
        self.id = 0
        self.PathImport = ""
        self.df1 = pd.DataFrame()
        self.df2 = pd.DataFrame()

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
        
    def Container_1(self, win, title):
        
        self.Frame_1 = tk.LabelFrame(
            win, 
            text=title, 
            font=("Helvetica 10 bold"), 
            fg="#004C8C", labelanchor='n')
        self.Frame_1.place(relx=0.02, rely=0.08, relheight=0.35, relwidth=0.96)

            
        self.VarLabelPath_1 = tk.StringVar()
        self.VarLabelPath_1.set("bla bla bla")
        self.LabelPath_1 = tk.Label(self.Frame_1, textvariable=self.VarLabelPath_1, bg="#FAEBD7")
        self.LabelPath_1.pack(fill="x")
        
        # Button import avec icon
        self.excelBtn_1 = tk.Button(
            self.Frame_1,
            image=self.excelIcon,
            text="Import data from Excel",
            compound="top",
            height=70,
            width=160,
            bd=1,
            bg="#DCDCDC",
            command=self.Excel,
            pady=2
        ).place(relx=0.23, rely=0.21)

        self.listbox_1 = tk.Listbox(self.Frame_1)
        self.listbox_1.place(relx=0.65, rely=0.17, relheight=0.8, relwidth=0.3)
        
        treescrolly = tk.Scrollbar(self.listbox_1, orient="vertical", command=self.listbox_1.yview)
        treescrollx = tk.Scrollbar(self.listbox_1, orient="horizontal", command=self.listbox_1.xview)
        self.listbox_1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
    
        self.ViewDataBtn_1 = tk.Button(
            self.Frame_1,
            image=self.viewIcon,
            text=f"   Voir le données {title}",
            compound="left",
            height=20,
            width=190,
            bd=1,
            bg="#DCDCDC",
            command=self.ViewData,
        ).place(relx=0.35, rely=0.8)
    
    def Container_2(self, win, title):
        
        self.Frame_2 = tk.LabelFrame(
            win, 
            text=title, 
            font=("Helvetica 10 bold"), 
            fg="#004C8C", labelanchor='n')
        self.Frame_2.place(relx=0.02, rely=0.45, relheight=0.35, relwidth=0.96)
            
        self.VarLabelPath_2 = tk.StringVar()
        self.VarLabelPath_2.set("bla bla bla")
        self.LabelPath_2 = tk.Label(self.Frame_2, textvariable=self.VarLabelPath_2, bg="#FAEBD7")
        self.LabelPath_2.pack(fill="x")
        
        # Button import avec icon
        self.excelBtn_2 = tk.Button(
            self.Frame_2,
            image=self.excelIcon,
            text="Import data from Excel",
            compound="top",
            height=70,
            width=160,
            bd=1,
            bg="#DCDCDC",
            command=self.Excel,
            pady=2
        ).place(relx=0.23, rely=0.21)

        self.listbox_2 = tk.Listbox(self.Frame_2)
        self.listbox_2.place(relx=0.65, rely=0.17, relheight=0.8, relwidth=0.3)
        
        treescrolly = tk.Scrollbar(self.listbox_2, orient="vertical", command=self.listbox_2.yview)
        treescrollx = tk.Scrollbar(self.listbox_2, orient="horizontal", command=self.listbox_2.xview)
        self.listbox_2.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
    
        self.ViewDataBtn_2 = tk.Button(
            self.Frame_2,
            image=self.viewIcon,
            text=f"   Voir le données {title}",
            compound="left",
            height=20,
            width=190,
            bd=1,
            bg="#DCDCDC",
            command=self.ViewData,
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
        
        self.Container_1(self.TopWindow, source)
        self.Container_2(self.TopWindow, 'Sharepoint')
        self.ContainerBtn(self.TopWindow)

    def Window_SAP_vs_Sharepoint(self):
        self.Toplevel_Window("SAP versus Sharepoint", "SAP")
        
    def Window_EuroDataHos_vs_Sharepoint(self):
        self.Toplevel_Window("EuroDataHos versus Sharepoint", "EuroDataHOS")

    def preview_data(self, path, df):
        """
        Cette sous fonction de la fonction Load_data_file() permet d'imorter les données et d'ouvrir une petite fenetre afin de prévisualiser les 5 premières lignes et enfin les données sont ok elle permet d'importer toutes les données
        """

        self.preview = tk.Toplevel(self.TopWindow)
        self.preview.title("Previous Data")
        self.preview.iconbitmap("media/TotalEnergies.ico")
        self.preview.geometry("600x250+15+15")
        self.preview.resizable(width=False, height=False)
        
        # Add Some Style
        style = ttk.Style()

        # Pick A Theme
        style.theme_use("clam")

        # Configure the Treeview Colors
        style.configure(
            "Treeview.Heading",
            background="lightblue",
            foreground="black",
            rowheight=25,
            fieldbackground="white",
        )
        # style.theme_use("clam")
        # style.configure(
        #     "Treeview.Heading", background="lightblue", foreground="black"
        # )

        # Change Selected Color
        style.map("Treeview", background=[("selected", "#347083")])

        def clear_data():
            self.tv_All_Data.delete(*self.tv_All_Data.get_children())
            return None

        # def ok_data_V():
        #     """Cette fonction valide les données et affiche toutes les données. Elle est relier au bouton ok pour valider"""

        #     self.fil_data_to_treeview_listbox(df, w="all")
        #     self.switchButtonState()
        #     self.preview.destroy()
        #     return df

        frame1 = tk.LabelFrame(self.preview, text=f"{path}")
        frame1.place(height=170, width=530, rely=0.02, relx=0.05)

        self.tv_All_Data = ttk.Treeview(frame1)
        self.tv_All_Data.place(relheight=1, relwidth=1)

        # commande signifie mettre à jour la vue de l'axe y du widget
        treescrolly = tk.Scrollbar(frame1, orient="vertical", command=self.tv_All_Data.yview)

        # commande signifie mettre à jour la vue axe x du widget
        treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=self.tv_All_Data.xview)

        # affecter les barres de défilement au widget Treeview
        self.tv_All_Data.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)

        # faire en sorte que la barre de défilement remplisse l'axe x du widget Treeview
        treescrollx.pack(side="bottom", fill="x")

        # faire en sorte que la barre de défilement remplisse l'axe y du widget Treeview
        treescrolly.pack(side="right", fill="y")

        # fram_check_btn_lbl = tk.Frame(self.preview)
        # fram_check_btn_lbl.place(relx=0.05, rely=0.73)

        # self.VarCheckBtn_add_index = tk.BooleanVar()
        # self.VarCheckBtn_add_index.set(True)
        # CheckBtn_add_index = tk.Checkbutton(
        #     fram_check_btn_lbl,
        #     variable=self.VarCheckBtn_add_index,
        #     command=None,
        # )
        # CheckBtn_add_index.grid(row=0, column=0)

        # text_checkbtn_add_index = tk.Label(
        #     fram_check_btn_lbl, text="Add an index column"
        # )
        # text_checkbtn_add_index.grid(row=0, column=1)

        OkBtn_data = tk.Button(
            self.preview,
            # text="Ok",
            # background="#40A497",
            # activeforeground="white",
            # activebackground="#40A497",
            text="OK",
            background="#6DA3F4",
            activebackground="#0256CD",
            foreground="white",
            activeforeground="white",
            width=12,
            height=1,
            command=None,
        ).place(relx=0.32, rely=0.87)

        Cancel_data = tk.Button(
            self.preview,
            text="Cancel",
            background="#CCCCCC",
            width=12,
            height=1,
            command=self.preview.destroy,
        ).place(relx=0.48, rely=0.87)

        global count
        count = 0

        self.tv_All_Data.tag_configure("oddrow", background="white")
        self.tv_All_Data.tag_configure("evenrow", background="#D3D3D3")

        # vider le treeview
        self.tv_All_Data.delete(*self.tv_All_Data.get_children())

        self.tv_All_Data["column"] = list(df.columns)
        self.tv_All_Data["show"] = "headings"

        for column in self.tv_All_Data["columns"]:
            self.tv_All_Data.column(column, anchor="w")
            self.tv_All_Data.heading(column, anchor="w", text=column)

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            if count % 2 == 0:
                self.tv_All_Data.insert(
                    "",
                    "end",
                    iid=count,
                    values=row,
                    tags=("evenrow",),
                )
            else:
                self.tv_All_Data.insert(
                    "",
                    "end",
                    iid=count,
                    values=row,
                    tags=("oddrow",),
                )
            count += 1

        self.tv_All_Data.insert("", "end", values="")

        return None
        
    def ImportData(self):

        """
        Cette grosse fonction permet d'abord d'ouvrir l'explorateur et parcourir le schéma du fichier, enssuite de le prévisualiser les 5 premières lignes et enfin les données sont ok elle permet d'importer toutes les données
        """

        if self.typefile == "Excel":
            path_filename = filedialog.askopenfilename(
                initialdir="E:\Total\Station Data\Master data\Data source",
                title="Select A File",
                filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")),
            )

        elif self.typefile == "CSV":
            path_filename = filedialog.askopenfilename(
                initialdir="E:\Total\Station Data\Master data\Data source",
                title="Select A File",
                filetype=(("csv files", "*.csv"), ("All Files", "*.*")),
            )

        else:
            path_filename = filedialog.askopenfilename(
                initialdir="E:\Total\Station Data\Master data\Data source",
                title="Select A File",
                filetype=(("All Files", "*.*")),
            )

        # print(path_filename[-4:])
        # print(path_filename)
        if path_filename:
            """Si le fichier sélectionné est valide, cela chargera le fichier"""

            try:
                df = EuroShare.LoadData(self,path_filename[-4:], path_filename)
            except ValueError:
                tk.messagebox.showerror("Information", "The file you have chosen is invalid")
                return None
            except FileNotFoundError:
                tk.messagebox.showerror("Information", f"No such file as {path_filename}")
                return None
            
            self.PathImport = path_filename
            if self.id == 1:
                self.VarLabelPath_1.set(path_filename)
                self.df1 = df
            elif self.id == 2:
                self.VarLabelPath_2.get(path_filename)
                self.df2 = df

        else:
            pass

    def ViewData(self):
        self.preview_data(self.PathImport, self.df2)

    def Excel(self):
        self.typefile = "Excel"
        self.ImportData()
    def CSV(self):
        self.typefile = "CSV"
        self.ImportData()
        
    def runing(self):
        pass
