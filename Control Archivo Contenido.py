import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory

# seleccionar una carpeta
ruta = askdirectory()

# crear una lista con los nombres de los archivos en la carpeta
archivos = os.listdir(ruta)

# crear un dataframe con los nombres de los archivos
df = pd.DataFrame(archivos, columns=['Archivo'])

# agregar la 'ruta' + '/' al 'Nompre'
df['Archivo con ruta'] = ruta + '/' + df['Archivo']

# Crear una columna que se llame primera celda
df['Primera celda'] = np.nan

# Por cada archivo en la columna 'Archivo' se debe leer el contnido de la primera celda y guardarlo en la columna 'Primera celda'
for i in range(len(df)):
    df['Primera celda'][i] = pd.read_excel(df['Archivo con ruta'][i], header=None).iloc[0,0]
del i

#Crear la columa 'CUIT Archivo' con el valor de archivo a partir de la posicion 20 con un largo de 11 caracteres
df['CUIT Archivo'] = df['Archivo'].str.slice(19, 30)

#Crear la columa 'CUIT Primera Celda' con el valor de los últimos 11 caracteres de la columna 'Primera celda'
df['CUIT Primera Celda'] = df['Primera celda'].str.slice(-11)

#Crear columna de 'Control' con el valor de 'SI' si los valores de las columnas 'CUIT Archivo' y 'CUIT Primera Celda' son iguales
df['Control'] = np.where(df['CUIT Archivo'] == df['CUIT Primera Celda'], 'SI', 'NO')

# Guardar el dataframe en un archivo excel
df.to_excel('Resultado.xlsx', index=False)