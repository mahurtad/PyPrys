import pandas as pd
import os
import re  # Biblioteca para usar expresiones regulares y eliminar caracteres no permitidos

# Ruta del archivo Excel original
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

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

# Filtrar registros que no tengan "Estado Entregable" igual a "Aprobado"
df_filtered = df_clean[(df_clean['Estado Entregable'] != 'Aprobado') | (df_clean['Estado Entregable'].isna())]

# Obtener la lista de PMs únicos en el filtro
unique_pms = df_filtered['PM Asignado'].unique()

# Ruta del archivo final
output_file = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\reporte_completo_pm.xlsx'

# Función para sanitizar nombres de hojas (reemplazar caracteres inválidos)
def sanitize_sheet_name(name):
    # Reemplazar caracteres no válidos por guiones
    return re.sub(r'[\\/*?:\[\]]', '-', name)

# Crear un archivo Excel con varias hojas
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Agregar la hoja principal "Plan Pruebas Integrales" con todos los registros filtrados
    df_clean[columns_to_select + ['Estado Entregable']].to_excel(writer, sheet_name='Plan Pruebas Integrales', index=False)
    
    # Crear una hoja por cada PM
    for pm in unique_pms:
        # Filtrar los datos del PM actual
        df_pm = df_filtered[df_filtered['PM Asignado'] == pm][columns_to_select]
        
        # Agregar la columna "Validación PM"
        df_pm['Validación PM'] = ""
        
        # Crear una hoja con el nombre del PM, sanitizando caracteres inválidos
        sanitized_pm_name = sanitize_sheet_name(pm)
        df_pm.to_excel(writer, sheet_name=sanitized_pm_name, index=False)
    
    print(f"Archivo exportado exitosamente en: {output_file}")
