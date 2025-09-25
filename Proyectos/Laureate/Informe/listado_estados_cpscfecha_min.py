import pandas as pd
import os

# Ruta del archivo Excel original
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\Calidad de Software - UPC\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 GL2 y GL3 v3.xlsx'

# Cargar el archivo Excel y la hoja correspondiente
df = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Asegurarse de que los nombres de las columnas estén limpios (sin espacios extra)
df.columns = df.columns.str.strip()

# Convertir las columnas de fechas al formato de fecha, ignorando errores
df['Fecha Inicial'] = pd.to_datetime(df['Fecha Inicial'], errors='coerce')
df['Fecha Planificada'] = pd.to_datetime(df['Fecha Planificada'], errors='coerce')

# Filtrar los casos de prueba que pertenecen a GL1, GL2 y GL3
df_gl1_gl2_gl3 = df[df['N° GL'].isin([1, 2, 3])]

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
        return 'Pendiente de Retest'

# Aplicar la lógica de agrupación por cada caso de prueba
df_reporte = df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP']).apply(determinar_estado_agrupado).reset_index(name='Estado de CP')

# Calcular la fecha más próxima en "Fecha Planificada"
fecha_planificada_min = df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP'])['Fecha Planificada'].min().reset_index(name='Fecha Planificada')

# Nueva función para concatenar ID y nombre de tareas
def formatear_tareas_estados_pm_con_id(sub_df):
    return '\n'.join(
        f"{id_d},{tarea} ({estado},{pm})"
        for id_d, tarea, estado, pm in zip(
            sub_df['N° D'].dropna().astype(str),
            sub_df['Nombre de Tarea (Desarrollo)'].dropna(),
            sub_df['Estado Entregable'].dropna(),
            sub_df['PM Asignado'].dropna()
        )
    )

# Generar el reporte con el formato actualizado
df_reporte = df_reporte.merge(
    df_gl1_gl2_gl3.groupby(['N° GL', 'N° CP']).apply(
        lambda x: pd.Series({
            'N° D': ', '.join(x['N° D'].dropna().astype(str).unique()),
            'Caso de Prueba': ', '.join(x['Caso de Prueba'].dropna().unique()),
            'Nombre de Tarea (Desarrollo)': formatear_tareas_estados_pm_con_id(x),
            'BPO': ', '.join(x['BPO'].dropna().unique()),
            'PM Asignado': ', '.join(x['PM Asignado'].dropna().unique()),
            'Periférico': ', '.join(x['Periférico'].dropna().unique())
        })
    ).reset_index(),
    on=['N° GL', 'N° CP']
)

# Incorporar la fecha planificada más próxima
df_reporte = df_reporte.merge(fecha_planificada_min, on=['N° GL', 'N° CP'])

# Seleccionar las columnas solicitadas en el orden deseado
columnas_seleccionadas = ['N° GL', 'N° CP', 'N° D', 'Caso de Prueba', 'Nombre de Tarea (Desarrollo)', 'BPO', 'PM Asignado', 'Periférico', 'Fecha Planificada', 'Estado de CP']
df_reporte = df_reporte[columnas_seleccionadas]

# Exportar los resultados a un archivo Excel
output_path = os.path.join(os.path.dirname(file_path_new), 'lista_estado_cps_min_04032025.xlsx')
df_reporte.to_excel(output_path, index=False)

print(f"El archivo ha sido guardado en: {output_path}")
