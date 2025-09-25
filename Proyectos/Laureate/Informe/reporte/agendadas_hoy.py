import pandas as pd
import os

# Ruta del archivo actualizado
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar el archivo Excel y la hoja correspondiente
plan_pruebas_integrales_df_new = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Convertir la columna Fecha Planificada a formato de fecha, ignorando errores
plan_pruebas_integrales_df_new['Fecha Planificada'] = pd.to_datetime(plan_pruebas_integrales_df_new['Fecha Planificada'], format='%d/%m/%Y', errors='coerce')

# Definir la fecha de corte (modificar según se necesite)
fecha_corte_actual = pd.to_datetime('19/10/2024', format='%d/%m/%Y')

# Filtrar el dataframe por la fecha de corte ingresada
df_filtrado = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['Fecha Planificada'] <= fecha_corte_actual]

# Seleccionar las columnas requeridas
columnas_requeridas = ['N° GL', 'N° CP', 'N° D', 'Caso de Prueba', 'Nombre de Tarea (Desarrollo)', 'Fecha Planificada',
                       'BPO', 'PM Asignado', 'Periférico', 'Estado Entregable']

df_final = df_filtrado[columnas_requeridas]

# Definir la ruta de salida para el archivo Excel
output_path = os.path.join(os.path.dirname(file_path_new), 'reporte_casos_prueba.xlsx')

# Exportar los resultados a Excel
df_final.to_excel(output_path, index=False)

# Mostrar los resultados en consola
print(df_final)

# Mostrar la ruta del archivo exportado
print(f"El archivo ha sido guardado en: {output_path}")
