import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory
import openpyxl

# seleccionar una carpeta
ruta = askdirectory()

# Crear una lista con los nombres de los archivos en la carpeta
archivos = os.listdir(ruta)

# Filtrar los archivos que contengan 'MCR'
archivos = [archivo for archivo in archivos if 'MCR' in archivo]

# Crear un dataframe con los nombres de los archivos
Tabla_Archivos = pd.DataFrame(archivos, columns=['Archivo'])

# Agregar la 'ruta' + '/' al 'Nompre'
Tabla_Archivos['Archivo con ruta'] = ruta + '/' + Tabla_Archivos['Archivo']

# Crear una columna que se llame primera celda
Tabla_Archivos['Primera celda'] = np.nan

# Leer cada archivo y guardar el contenido de la primera celda en una columna llamada 'Primera celda'
for i in range(len(Tabla_Archivos)):
    Tabla_Archivos['Primera celda'][i] = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], header=None).iloc[0,0]
    del i

#crear la Carpeta 'Compras Procesadas' en la 'ruta' si no existe
if not os.path.exists(ruta + '/Compras Procesadas'):
    os.makedirs(ruta + '/Compras Procesadas')

#Leer cada Excel en un Dataframe_Temporal y Exportarlo a la carpeta 'Compras Procesadas'
for i in range(len(Tabla_Archivos)):
    df = pd.read_excel(Tabla_Archivos['Archivo con ruta'][i], skiprows=1)
    # No filtrar los datos donde el 'Tipo' contenga ' B'
    df = df[df['Tipo'].str.contains(' B') == False]
    df.to_excel(ruta + '/Compras Procesadas/' + "Procesado - " + Tabla_Archivos['Archivo'][i], index=False)
    
    #Al archivo recien creado, agregar la primera celda en la primera fila
    wb = openpyxl.load_workbook(ruta + '/Compras Procesadas/' + "Procesado - " + Tabla_Archivos['Archivo'][i])
    ws = wb.active
    ws.insert_rows(1)
    ws['A1'] = Tabla_Archivos['Primera celda'][i]
    wb.save(ruta + '/Compras Procesadas/' + "Procesado - " + Tabla_Archivos['Archivo'][i])
    
    del i, df, wb, ws


# # Consolidar todos los archivos de la carpeta 'Compras Procesadas' en un solo archivo
# archivos = os.listdir(ruta + '/Compras Procesadas')
# archivos = [archivo for archivo in archivos if 'Procesado' in archivo]
# df = pd.DataFrame()
# for i in range(len(archivos)):
#     df_temp = pd.read_excel(ruta + '/Compras Procesadas/' + archivos[i] , skiprows=1)
#     df = pd.concat([df, df_temp], axis=0)
#     del i, df_temp
# df.to_excel(ruta + '/Compras Procesadas/Consolidado.xlsx', index=False)
