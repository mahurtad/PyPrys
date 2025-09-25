import pandas as pd
import os

# Ruta del archivo Excel original
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar la hoja "Plan Pruebas Integrales"
df = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales')

# Renombrar columnas para que coincidan con los nombres esperados
df_clean = df.rename(columns={
    'N° GL': 'Numero GL',
    'N° CP': 'Numero CP',
    'N° D': 'Numero D',  # Incluimos la columna N° D
    'Caso de Prueba': 'Caso de Prueba',
    'Fecha Planificada': 'Fecha Planificada',
    'PM Asignado': 'PM Asignado',
    'Periférico': 'Periférico',
    'Estado Entregable': 'Estado Entregable',
    'Nombre de Tarea (Desarrollo)': 'Nombre de Tarea Desarrollo'  # Incluimos la columna Nombre de Tarea
})

# Definir las columnas que deseas extraer, colocando 'Nombre de Tarea Desarrollo' después de 'Caso de Prueba'
columns_to_select = ['Numero GL', 'Numero CP', 'Numero D', 'Caso de Prueba', 'Nombre de Tarea Desarrollo', 'Fecha Planificada', 'PM Asignado', 'Periférico', 'Estado Entregable']

# Filtrar por registros que contengan "Karla Zayerz" o "Karla Zayerz, Julio Rodriguez" en el campo "PM Asignado"
df_filtered = df_clean[df_clean['PM Asignado'].str.contains("Karla Zayerz|Karla Zayerz,Julio Rodriguez", na=False, case=False)]

# Extraer todos los registros que correspondan a los mismos "Numero CP" encontrados
df_final = df_clean[df_clean['Numero CP'].isin(df_filtered['Numero CP'])][columns_to_select]

# Obtener la ruta del directorio de donde se cargó el archivo
output_dir = os.path.dirname(file_path)

# Crear la ruta para el archivo de salida
output_file = os.path.join(output_dir, 'pruebas_integrales_filtradas.xlsx')

# Guardar los resultados en un archivo Excel en la misma ruta
df_final.to_excel(output_file, index=False)

print(f"Archivo exportado exitosamente en: {output_file}")
