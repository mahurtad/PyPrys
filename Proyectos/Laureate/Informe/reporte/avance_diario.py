import pandas as pd
from tabulate import tabulate

# Ruta del archivo Excel original
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar la hoja "Plan Pruebas Integrales"
df = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales')

# Renombrar columnas para facilitar el acceso
df_clean = df.rename(columns={
    'N° GL': 'Numero GL',
    'N° CP': 'Numero CP',
    'Fecha Planificada': 'Fecha Planificada',
    'Estado Entregable': 'Estado Entregable'
})

# Asegurarnos que la columna de fechas esté en formato datetime
df_clean['Fecha Planificada'] = pd.to_datetime(df_clean['Fecha Planificada'], errors='coerce').dt.date

# Filtramos sólo las columnas que vamos a usar
df_filtered = df_clean[['Numero GL', 'Numero CP', 'Fecha Planificada', 'Estado Entregable']]

# Función para contar los estados de Planificados, Aprobados y Retest
def contar_estado(df, estado, gl, fecha):
    if estado == 'Aprobadas':
        aprobados = df[(df['Numero GL'] == gl) & (df['Fecha Planificada'] == fecha)]
        # Filtrar sólo los casos donde todos los registros del mismo CP estén aprobados
        casos_aprobados = aprobados.groupby('Numero CP').filter(lambda x: (x['Estado Entregable'] == 'Aprobado').all())
        return casos_aprobados['Numero CP'].nunique(), casos_aprobados['Numero CP'].unique()
    elif estado == 'Retest':
        retest = df[(df['Numero GL'] == gl) & (df['Estado Entregable'] == 'Pendiente retest') & (df['Fecha Planificada'] == fecha)]
        return retest['Numero CP'].nunique(), retest['Numero CP'].unique()
    else:
        planificados = df[(df['Numero GL'] == gl) & (df['Fecha Planificada'] == fecha)]
        return planificados['Numero CP'].nunique(), planificados['Numero CP'].unique()

# Función para obtener los códigos de los casos de prueba cancelados
def obtener_casos_cancelados(df, gl, fecha):
    planificados = df[(df['Numero GL'] == gl) & (df['Fecha Planificada'] == fecha)]
    no_cancelados = planificados[planificados['Estado Entregable'].isin(['Aprobado', 'Pendiente retest'])]['Numero CP']
    cancelados = planificados[~planificados['Numero CP'].isin(no_cancelados)]['Numero CP'].unique()
    return ', '.join(map(str, cancelados)) if len(cancelados) > 0 else "N/A"

# Definir una fecha específica para el reporte
fecha_especifica = pd.to_datetime('2024-10-23').date()  # Cambia esta fecha según sea necesario

# Crear el reporte para una fecha específica
planificadas_gl1, cp_planificadas_gl1 = contar_estado(df_filtered, 'Planificados', 1, fecha_especifica)
planificadas_gl3, cp_planificadas_gl3 = contar_estado(df_filtered, 'Planificados', 3, fecha_especifica)

aprobadas_gl1, cp_aprobadas_gl1 = contar_estado(df_filtered, 'Aprobadas', 1, fecha_especifica)
aprobadas_gl3, cp_aprobadas_gl3 = contar_estado(df_filtered, 'Aprobadas', 3, fecha_especifica)

retest_gl1, cp_retest_gl1 = contar_estado(df_filtered, 'Retest', 1, fecha_especifica)
retest_gl3, cp_retest_gl3 = contar_estado(df_filtered, 'Retest', 3, fecha_especifica)

# Calcular las cancelaciones (Planificados - Aprobados - Retest)
cancelados_gl1 = planificadas_gl1 - aprobadas_gl1 - retest_gl1
cancelados_gl3 = planificadas_gl3 - aprobadas_gl3 - retest_gl3

# Obtener los códigos de los casos de prueba cancelados
cp_cancelados_gl1 = obtener_casos_cancelados(df_filtered, 1, fecha_especifica)
cp_cancelados_gl3 = obtener_casos_cancelados(df_filtered, 3, fecha_especifica)

# Crear un DataFrame para exportar los resultados
reporte = pd.DataFrame({
    'Fecha': [fecha_especifica],
    'Planificadas_GL1': [planificadas_gl1],
    'Planificadas_GL3': [planificadas_gl3],
    'Aprobadas_GL1': [aprobadas_gl1],
    'Aprobadas_GL3': [aprobadas_gl3],
    'Retest_GL1': [retest_gl1],
    'Retest_GL3': [retest_gl3],
    'Canceladas_GL1': [cancelados_gl1],
    'Canceladas_GL3': [cancelados_gl3],
    'CP Canceladas_GL1': [cp_cancelados_gl1],
    'CP Canceladas_GL3': [cp_cancelados_gl3]
})

# Crear una segunda fila con los códigos de los casos de prueba
reporte_cps = pd.DataFrame({
    'Fecha': ['Casos:'],
    'Planificadas_GL1': [', '.join(map(str, cp_planificadas_gl1))],
    'Planificadas_GL3': [', '.join(map(str, cp_planificadas_gl3))],
    'Aprobadas_GL1': [', '.join(map(str, cp_aprobadas_gl1))],
    'Aprobadas_GL3': [', '.join(map(str, cp_aprobadas_gl3))],
    'Retest_GL1': [', '.join(map(str, cp_retest_gl1))],
    'Retest_GL3': [', '.join(map(str, cp_retest_gl3))],
    'Canceladas_GL1': [cp_cancelados_gl1],
    'Canceladas_GL3': [cp_cancelados_gl3]
})

# Concatenar ambos DataFrames
reporte_final = pd.concat([reporte, reporte_cps], ignore_index=True)

# Crear el nombre del archivo usando la fecha de corte
output_file = rf'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\reporte_resumen_{fecha_especifica}_.xlsx'

# Guardar el reporte en un archivo Excel con formato usando XlsxWriter
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    reporte_final.to_excel(writer, sheet_name='Resumen', index=False)

    # Obtener el objeto workbook y worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Resumen']

    # Definir formato para los encabezados (color de fondo verde claro, negrita y alineación)
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'center',
        'fg_color': '#C6EFCE',  # Color verde claro
        'border': 1
    })

    # Aplicar el formato a los encabezados
    for col_num, value in enumerate(reporte.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Ajustar el ancho de las columnas para mejor visualización
    worksheet.set_column(0, len(reporte.columns) - 1, 20)

print(f"Reporte generado y guardado exitosamente en {output_file}")
