import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory, askopenfilename
import openpyxl


# seleccionar una archivo
ruta = askopenfilename(title="Seleccionar el archivo que contiene los archivos a procesar")

# Crear una lista con los nombres de los archivos en la carpeta
ListaArchivos = pd.read_csv(ruta, header=None, sep="	")[1].tolist()
del ruta

# Eliminar el primer elemento de la lista
del ListaArchivos[0]

# Reemplazar los '\\' por '/' en cada elemento de la lista
ListaArchivos = [archivo.replace('\\', '/') for archivo in ListaArchivos]

ListaArchivosParaIterar = []

# Listar todos los archivos de cada elemento de la lista 'ListaArchivos' en una lista con su ruta completa
for i in ListaArchivos:
    Archivos = os.listdir(i)
    Archivos = [i + '/' + archivo for archivo in Archivos]
    # Agregar a la lista ArchivosParaIterar los elementos de la lista Archivos
    ListaArchivosParaIterar.extend(Archivos)
    del i , Archivos

ListaArchivos = ListaArchivosParaIterar
del ListaArchivosParaIterar

# Filtrar los que no son '.xlsx'
archivos = [archivo for archivo in ListaArchivos if '.xlsx' in archivo]

# Quitar los archivos que empiezan con '~$'
archivos = [archivo for archivo in archivos if '~$' not in archivo]

# Crear un dataframe con los nombres de los archivos
Tabla_Archivos = pd.DataFrame(archivos, columns=['Archivo'])
del archivos

# Crear una columna que se llame primera celda
Tabla_Archivos['Primera celda'] = np.nan

# Crear una columna que se llame 'ruta' y que contenga la columna 'Archivo' sin el nombre del archivo
Tabla_Archivos['ruta'] = Tabla_Archivos['Archivo'].str.replace(r'[^/]+$', '' , regex=True)

def obtener_primera_celda(archivo):
    return pd.read_excel(archivo, header=None).iloc[0,0]

Tabla_Archivos['Primera celda'] = Tabla_Archivos['Archivo'].apply(obtener_primera_celda)

# Crear la Carpeta 'Procesado' en la 'ruta' si no existe de la columna 'ruta' de Tabla_Archivos
for i in range(len(Tabla_Archivos)):
    if not os.path.exists(Tabla_Archivos['ruta'][i] + '/Procesado'):
        os.makedirs(Tabla_Archivos['ruta'][i] + '/Procesado')
    del i

#Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Procesado'
for i in range(len(Tabla_Archivos)):
    df = pd.read_excel(Tabla_Archivos['Archivo'][i], skiprows=1)
    # si la columna 'Archivo' contiene ' MCR ' entonces filtrar los datos donde el 'Tipo' contenga ' B'
    if ' MCR ' in Tabla_Archivos['Archivo'][i]:
        df = df[df['Tipo'].str.contains(' B') == False]
    #Exportar el dataframe solo si contiene datos
    if len(df) > 0:
        #Si las columnas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' estan vacias entonces rellenarlas con 0
        df['Imp. Neto Gravado'] = df['Imp. Neto Gravado'].fillna(0)
        df['Imp. Neto No Gravado'] = df['Imp. Neto No Gravado'].fillna(0)
        df['Imp. Op. Exentas'] = df['Imp. Op. Exentas'].fillna(0)
        df['IVA'] = df['IVA'].fillna(0)
        df['Imp. Total'] = df['Imp. Total'].fillna(0)

        #Si la columan 'Tipo' contiene 'Nota de Crédito' entonces multimplicar por -1 las columnas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'
        df.loc[df["Tipo"].str.contains("Nota de Crédito"), ['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']] *= -1

        # Multiplicar las columnas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total' por la columna 'Tipo Cambio'
        df['Imp. Neto Gravado'] *= df['Tipo Cambio']
        df['Imp. Neto No Gravado'] *= df['Tipo Cambio']
        df['Imp. Op. Exentas'] *= df['Tipo Cambio']
        df['IVA'] *= df['Tipo Cambio']
        df['Imp. Total'] *= df['Tipo Cambio']

        # Agregar una fila con la suma de las columnas 'Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total'
        df = pd.concat([df, pd.DataFrame(df[['Imp. Neto Gravado' , 'Imp. Neto No Gravado' , 'Imp. Op. Exentas' , 'IVA' , 'Imp. Total']].sum()).T], ignore_index=True)

        # Exportar el dataframe a la carpeta 'Procesado' con el nombre del archivo + ' - Procesado.xlsx'
        df.to_excel(Tabla_Archivos['ruta'][i] + '/Procesado/' + Tabla_Archivos['Archivo'][i].split('/')[-1] + " - Procesado.xlsx", index=False)

        #Al archivo recien creado, agregar la primera celda en la primera fila
        wb = openpyxl.load_workbook(Tabla_Archivos['ruta'][i] + '/Procesado/' + Tabla_Archivos['Archivo'][i].split('/')[-1] + " - Procesado.xlsx")
        ws = wb.active
        ws.insert_rows(1)
        ws['A1'] = Tabla_Archivos['Primera celda'][i]
        wb.save(Tabla_Archivos['ruta'][i] + '/Procesado/' + Tabla_Archivos['Archivo'][i].split('/')[-1] + " - Procesado.xlsx")
    
del i, df, wb, ws
