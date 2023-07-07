import pandas as pd
import os
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo

# Preguntar por el directorio
directorio = askdirectory(title = "Selecciona la carpeta con los archivos Excel")

# lista de archivos en el directorio y filtrar los que no son .xlsx
lista_archivos = os.listdir(directorio)
lista_archivos = [archivo for archivo in lista_archivos if archivo.endswith(".xlsx")]

# Crear una carpeta "CSV" en el directorio
if not os.path.exists(directorio + "/CSV"):
    os.makedirs(directorio + "/CSV")

# Por cada archivo en la lista de archivos leer el excel salteando la primer fila y exportar a CSV con el mismo nombre
for archivo in lista_archivos:
    df = pd.read_excel(directorio + "/" + archivo, skiprows=1)
    # Exportar el dataframe solo si contiene datos
    if len(df) > 0:
        df.to_csv(directorio + "/CSV/" + archivo[:-5] + ".csv", index=False)

# Mostrar mensaje de finalizado
showinfo("Finalizado", "Se han exportado los archivos a CSV en la carpeta CSV")

