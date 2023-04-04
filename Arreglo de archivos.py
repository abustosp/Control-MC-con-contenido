import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory
import openpyxl

# seleccionar una carpeta
ruta = askdirectory()

# Crear una lista con los nombres de los archivos en la carpeta
archivos = os.listdir(ruta)

# Filtrar los archivos que contengan ' MCR '
archivosMCR = [archivo for archivo in archivos if ' MCR ' in archivo]

# Crear un dataframe con los nombres de los archivos
Tabla_ArchivosMCR = pd.DataFrame(archivosMCR, columns=['Archivo'])

# Agregar la 'ruta' + '/' al 'Nompre'
Tabla_ArchivosMCR['Archivo con ruta'] = ruta + '/' + Tabla_ArchivosMCR['Archivo']

# Crear una columna que se llame primera celda
Tabla_ArchivosMCR['Primera celda'] = np.nan

#Filtrar los archivos que contengan ' MCE '
archivosMCE = [archivo for archivo in archivos if ' MCE ' in archivo]

# Crear un dataframe con los nombres de los archivos
Tabla_ArchivosMCE = pd.DataFrame(archivosMCE, columns=['Archivo'])

# Agregar la 'ruta' + '/' al 'Nompre'
Tabla_ArchivosMCE['Archivo con ruta'] = ruta + '/' + Tabla_ArchivosMCE['Archivo']

# Crear una columna que se llame primera celda
Tabla_ArchivosMCE['Primera celda'] = np.nan


# Leer cada archivo y guardar el contenido de la primera celda en una columna llamada 'Primera celda'
# for i in range(len(Tabla_Archivos)):
#     Tabla_Archivos['Primera celda'][i] = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], header=None).iloc[0,0]
#     del i

def obtener_primera_celda(archivo):
    return pd.read_excel(archivo, header=None).iloc[0,0]

Tabla_ArchivosMCR['Primera celda'] = Tabla_ArchivosMCR['Archivo con ruta'].apply(obtener_primera_celda)
Tabla_ArchivosMCE['Primera celda'] = Tabla_ArchivosMCE['Archivo con ruta'].apply(obtener_primera_celda)

#crear la Carpeta 'Procesado' en la 'ruta' si no existe
if not os.path.exists(ruta + '/Procesado'):
    os.makedirs(ruta + '/Procesado')

#Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Procesado'
for i in range(len(Tabla_ArchivosMCR)):
    df = pd.read_excel(Tabla_ArchivosMCR['Archivo con ruta'][i], skiprows=1)
    # No filtrar los datos donde el 'Tipo' contenga ' B'
    df = df[df['Tipo'].str.contains(' B') == False]
    df.to_excel(ruta + '/Procesado/' + "Procesado - " + Tabla_ArchivosMCR['Archivo'][i], index=False)
    
    #Al archivo recien creado, agregar la primera celda en la primera fila
    wb = openpyxl.load_workbook(ruta + '/Procesado/' + "Procesado - " + Tabla_ArchivosMCR['Archivo'][i])
    ws = wb.active
    ws.insert_rows(1)
    ws['A1'] = Tabla_ArchivosMCR['Primera celda'][i]
    wb.save(ruta + '/Procesado/' + "Procesado - " + Tabla_ArchivosMCR['Archivo'][i])
    
    del i, df, wb, ws

#Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Procesado'
for i in range(len(Tabla_ArchivosMCE)):
    df = pd.read_excel(Tabla_ArchivosMCE['Archivo con ruta'][i], skiprows=1)
    df.to_excel(ruta + '/Procesado/' + Tabla_ArchivosMCE['Archivo'][i], index=False)
    
    #Al archivo recien creado, agregar la primera celda en la primera fila
    wb = openpyxl.load_workbook(ruta + '/Procesado/' + Tabla_ArchivosMCE['Archivo'][i])
    ws = wb.active
    ws.insert_rows(1)
    ws['A1'] = Tabla_ArchivosMCE['Primera celda'][i]
    wb.save(ruta + '/Procesado/' + Tabla_ArchivosMCE['Archivo'][i])
    
    del i, df, wb, ws

# # Consolidar todos los archivos de la carpeta 'Procesado' en un solo archivo
# archivos = os.listdir(ruta + '/Procesado')
# archivos = [archivo for archivo in archivos if 'Procesado' in archivo]
# df = pd.DataFrame()
# for i in range(len(archivos)):
#     df_temp = pd.read_excel(ruta + '/Procesado/' + archivos[i] , skiprows=1)
#     df = pd.concat([df, df_temp], axis=0)
#     del i, df_temp
# df.to_excel(ruta + '/Procesado/Consolidado.xlsx', index=False)

