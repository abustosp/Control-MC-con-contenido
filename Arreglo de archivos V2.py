import os
import pandas as pd
import numpy as np
from tkinter.filedialog import askdirectory
import openpyxl

# seleccionar una carpeta
ruta = askdirectory()

# Crear una lista con los nombres de los archivos en la carpeta
archivos = os.listdir(ruta)

# Filtrar los que no son '.xlsx'
archivos = [archivo for archivo in archivos if '.xlsx' in archivo]

# Quitar los archivos que empiezan con '~$'
archivos = [archivo for archivo in archivos if '~$' not in archivo]

# reemplazar '.xlsx' por ''
archivos = [archivo.replace('.xlsx', '') for archivo in archivos]

# Crear un dataframe con los nombres de los archivos
Tabla_Archivos = pd.DataFrame(archivos, columns=['Archivo'])

# Agregar la 'ruta' + '/' al 'Nompre'
Tabla_Archivos['Archivo con ruta'] = ruta + '/' + Tabla_Archivos['Archivo'] + '.xlsx'

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

        df.to_excel(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx", index=False)


    if len(df) > 0:
        #Al archivo recien creado, agregar la primera celda en la primera fila
        wb = openpyxl.load_workbook(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx")
        ws = wb.active
        ws.insert_rows(1)
        ws['A1'] = Tabla_Archivos['Primera celda'][i]
        wb.save(ruta + '/Procesado/' + Tabla_Archivos['Archivo'][i] + " - Procesado.xlsx")
    
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

