import pandas as pd
import os
import numpy as np

# Ruta del archivo Excel original
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Ruta de la carpeta donde se guardarán los reportes por PM
output_dir = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\reporte_pms'

# Cargar la hoja "Plan Pruebas Integrales"
df = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales')

# Renombrar columnas para que coincidan con los nombres esperados
df_clean = df.rename(columns={
    'N° GL': 'Numero GL',
    'N° CP': 'Numero CP',
    'N° D': 'Numero D',  
    'Caso de Prueba': 'Caso de Prueba',
    'Fecha Planificada': 'Fecha Planificada',
    'PM Asignado': 'PM Asignado',
    'Periférico': 'Periférico',
    'Estado Entregable': 'Estado Entregable',
    'Nombre de Tarea (Desarrollo)': 'Nombre de Tarea Desarrollo',
    'Hora de Inicio': 'Hora de Inicio',
    'Hora de Fin': 'Hora de Fin',
    'BPO': 'BPO'  # Incluimos la columna BPO
})

# Formatear la columna 'Fecha Planificada' al formato DD/MM/AAAA
df_clean['Fecha Planificada'] = pd.to_datetime(df_clean['Fecha Planificada'], errors='coerce').dt.strftime('%d/%m/%Y')

# Definir las columnas que deseas extraer, incluyendo 'BPO' después de 'Fecha Planificada'
columns_to_select = [
    'Numero GL', 'Numero CP', 'Numero D', 'Caso de Prueba', 
    'Nombre de Tarea Desarrollo', 'Fecha Planificada', 'BPO',
    'PM Asignado', 'Periférico', 'Estado Entregable', 
    'Hora de Inicio', 'Hora de Fin'
]

# Filtrar registros que no tengan "Estado Entregable" igual a "Aprobado" o estén vacíos (NaN)
df_filtered = df_clean[(df_clean['Estado Entregable'] != 'Aprobado') | (df_clean['Estado Entregable'].isna())]

# Obtener la lista de PMs únicos en el filtro
unique_pms = df_filtered['PM Asignado'].unique()

# Crear la carpeta de salida si no existe
os.makedirs(output_dir, exist_ok=True)

# Generar un archivo Excel por cada PM
for pm in unique_pms:
    # Filtrar por el PM actual
    df_pm = df_filtered[df_filtered['PM Asignado'] == pm][columns_to_select]
    
    # Generar el nombre del archivo basado en el nombre del PM
    sanitized_pm_name = pm.replace(" ", "_").replace(",", "_").replace("\n", "_")
    output_file = os.path.join(output_dir, f'reporte_{sanitized_pm_name}.xlsx')
    
    # Guardar el archivo Excel para el PM actual
    df_pm.to_excel(output_file, index=False)
    
    print(f"Archivo exportado para {pm}: {output_file}")

print("Archivos exportados exitosamente.")
