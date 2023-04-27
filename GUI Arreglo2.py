import tkinter as tk
import tkinter.ttk as ttk
import LIB.ArregloV2 as ArregloV2


class App_Arreglo_MC:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=275,
            width=275)
        Toplevel_1.iconbitmap("LIB/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.minsize(275, 250)
        Toplevel_1.overrideredirect("False")
        Toplevel_1.title("Arreglo de Archivos de MC")
        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="LIB/ABP blanco en fondo negro111.png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            font="TkDefaultFont",
            foreground="#ffffff",
            justify="right",
            takefocus=True,
            text='Arreglo de Archivos de Mis Comprobantes de manera masiva',
            wraplength=325)
        Label_1.pack(expand="true", side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand="true", side="top")
        self.Mensual_XLS = ttk.Button(Toplevel_1)
        self.Mensual_XLS.configure(text='Seleccionar Archivo Excel con directorios' , command=ArregloV2.ArregloV2)
        self.Mensual_XLS.pack(expand="true", pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = App_Arreglo_MC()
    app.run()
