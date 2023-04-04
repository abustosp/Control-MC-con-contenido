import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory
import openpyxl

# seleccionar una carpeta
ruta = askdirectory()

# Crear una lista con los nombres de los archivos en la carpeta
archivos = os.listdir(ruta)

# Crear un dataframe con los nombres de los archivos
Tabla_Archivos = pd.DataFrame(archivos, columns=['Archivo'])

# Agregar la 'ruta' + '/' al 'Nompre'
Tabla_Archivos['Archivo con ruta'] = ruta + '/' + Tabla_Archivos['Archivo']

# Crear una columna que se llame primera celda
Tabla_Archivos['Primera celda'] = np.nan

# Leer cada archivo y guardar el contenido de la primera celda en una columna llamada 'Primera celda'
# for i in range(len(Tabla_Archivos)):
#     Tabla_Archivos['Primera celda'][i] = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], header=None).iloc[0,0]
#     del i

def obtener_primera_celda(archivo):
    return pd.read_excel(archivo, header=None).iloc[0,0]

Tabla_Archivos['Primera celda'] = Tabla_Archivos['Archivo con ruta'].apply(obtener_primera_celda)

#crear la Carpeta 'Procesado' en la 'ruta' si no existe
if not os.path.exists(ruta + '/Procesado'):
    os.makedirs(ruta + '/Procesado')

#Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Procesado'
for i in range(len(Tabla_Archivos)):
    df = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], skiprows=1)
    # si la columna 'Archivo' contiene ' MCR ' entonces filtrar los datos donde el 'Tipo' contenga ' B'
    if ' MCR ' in Tabla_Archivos['Archivo'][i]:
        df = df[df['Tipo'].str.contains(' B') == False]
        df.to_excel(ruta + '/Procesado/' + "Procesado - " + Tabla_Archivos['Archivo'][i], index=False)
    else:
        df.to_excel(ruta + '/Procesado/' + "Procesado - " + Tabla_Archivos['Archivo'][i], index=False)

    #Al archivo recien creado, agregar la primera celda en la primera fila
    wb = openpyxl.load_workbook(ruta + '/Procesado/' + "Procesado - " + Tabla_Archivos['Archivo'][i])
    ws = wb.active
    ws.insert_rows(1)
    ws['A1'] = Tabla_Archivos['Primera celda'][i]
    wb.save(ruta + '/Procesado/' + "Procesado - " + Tabla_Archivos['Archivo'][i])
    
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

