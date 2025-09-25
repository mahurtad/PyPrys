import pandas as pd
import os

# Ruta del archivo actualizado
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\Calidad de Software - UPC\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 GL2 y GL3 v3.xlsx'

# Cargar el archivo Excel y la hoja correspondiente
df = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Convertir las columnas de fechas a formato de fecha, ignorando errores
df['Fecha Inicial'] = pd.to_datetime(df['Fecha Inicial'], format='%d/%m/%Y', errors='coerce')
df['Fecha Planificada'] = pd.to_datetime(df['Fecha Planificada'], format='%d/%m/%Y', errors='coerce')

# Filtrar los casos de prueba que pertenecen a GL1, GL2 y GL3
df_gl1_gl2_gl3 = df[(df['N° GL'] == 1) | (df['N° GL'] == 2) | (df['N° GL'] == 3)]

# Definir la función para asignar el "Estado de CP"
def asignar_estado(row):
    if row['Estado Entregable'] == 'Aprobado':
        return 'Aprobado'
    elif row['Estado Entregable'] == 'No ejecutado':
        return 'No ejecutado'
    else:
        return 'Pendiente de Retest'

# Aplicar la función de estado
df_gl1_gl2_gl3['Estado de CP'] = df_gl1_gl2_gl3.apply(asignar_estado, axis=1)

# Agrupar por "N° GL" y "N° CP" para verificar que todos los registros cumplen con el estado "Aprobado" o "No ejecutado"
def determinar_estado_agrupado(sub_df):
    if all(sub_df['Estado de CP'] == 'Aprobado'):
        return 'Aprobado'
    elif all(sub_df['Estado de CP'] == 'No ejecutado'):
        return 'No ejecutado'
    else:
        return 'Pendiente retest/En progreso'

# Aplicar la lógica de agrupación por cada caso de prueba
df_reporte = df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP']).apply(determinar_estado_agrupado).reset_index(name='Estado de CP')

# Calcular la fecha más lejana en "Fecha Planificada"
fecha_planificada_max = df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP'])['Fecha Planificada'].max().reset_index(name='Fecha Planificada')

# Concatenar valores únicos de otros campos para el reporte final
df_reporte = df_reporte.merge(
    df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP']).agg({
        'N° D': lambda x: ', '.join(x.dropna().astype(str).unique()),
        'Caso de Prueba': lambda x: ', '.join(x.dropna().unique()),
        'Nombre de Tarea (Desarrollo)': lambda x: ', '.join(x.dropna().unique()),
        'BPO': lambda x: ', '.join(x.dropna().unique()),
        'PM Asignado': lambda x: ', '.join(x.dropna().unique()),
        'Periférico': lambda x: ', '.join(x.dropna().unique())
    }).reset_index(),
    on=['N° GL', 'N° CP']
)

# Incorporar la fecha planificada más lejana
df_reporte = df_reporte.merge(fecha_planificada_max, on=['N° GL', 'N° CP'])

# Seleccionar las columnas solicitadas en el orden deseado
columnas_seleccionadas = ['N° GL', 'N° CP', 'N° D', 'Caso de Prueba', 'Nombre de Tarea (Desarrollo)', 'BPO', 'PM Asignado', 'Periférico', 'Fecha Planificada','Estado de CP']
df_reporte = df_reporte[columnas_seleccionadas]

# Exportar los resultados a un archivo Excel
output_path = os.path.join(os.path.dirname(file_path_new), 'matriz_estados_cps.xlsx')
df_reporte.to_excel(output_path, index=False)

print(f"El archivo ha sido guardado en: {output_path}")
