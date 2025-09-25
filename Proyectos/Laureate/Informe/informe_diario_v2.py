import pandas as pd
from tabulate import tabulate

# Ruta del archivo actualizado
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar el archivo Excel y la hoja correspondiente
plan_pruebas_integrales_df_new = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Convertir la columna Fecha Planificada a formato de fecha, ignorando errores
plan_pruebas_integrales_df_new['Fecha Planificada'] = pd.to_datetime(plan_pruebas_integrales_df_new['Fecha Planificada'], format='%d/%m/%Y', errors='coerce')

# Función para centrar el contenido de las celdas
def centrar_contenido(celda):
    return f"{celda:^15}"

# Función para contar casos no ejecutados, donde todos los registros de "N° CP" están como "No ejecutado"
def contar_casos_no_ejecutados(df, gl_value):
    df_filtrado = df[df['N° GL'] == gl_value]
    no_ejecutados_df = df_filtrado.groupby('N° CP').filter(lambda x: (x['Estado Entregable'] == 'No ejecutado').all())
    return no_ejecutados_df['N° CP'].nunique()

# Función para contar los casos aprobados donde todos los registros de "N° CP" están aprobados
def contar_casos_aprobados(df, gl_value):
    df_filtrado = df[df['N° GL'] == gl_value]
    aprobados_df = df_filtrado.groupby('N° CP').filter(lambda x: (x['Estado Entregable'] == 'Aprobado').all())
    return aprobados_df['N° CP'].nunique()

# Función para calcular el porcentaje de casos ejecutados (Ejecutados / Total) * 100
def calcular_porcentaje_ejecutados(total_cases, ejecutados_cases):
    return "{:.2f}%".format((ejecutados_cases / total_cases) * 100) if total_cases > 0 else "0%"

# Función para calcular el porcentaje planificado
def calcular_porcentaje_planificado(df, gl_value):
    total_planificados = df[(df['N° GL'] == gl_value) & (df['Fecha Planificada'] <= fecha_limite)]['N° CP'].nunique()
    total_casos = df[df['N° GL'] == gl_value]['N° CP'].nunique()
    return "{:.2f}%".format((total_planificados / total_casos) * 100) if total_casos > 0 else "0%"

# Función para calcular la diferencia entre "% Casos Ejecutados" y "% Casos Ejecutados Planificado"
def calcular_diferencia(ejecutado, planificado):
    ejecutado_valor = float(ejecutado.replace('%', ''))
    planificado_valor = float(planificado.replace('%', ''))
    diferencia = ejecutado_valor - planificado_valor
    return f"{diferencia:.2f}%"

# Fecha límite para el cálculo (esta será la fecha de corte)
fecha_limite = pd.to_datetime('09/01/2025', format='%d/%m/%Y')

# Contar los casos de prueba únicos (N° CP) para GL1 y GL3
unique_cases_gl1_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 1]['N° CP'].nunique()
unique_cases_gl3_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 3]['N° CP'].nunique()

# Contar los casos no ejecutados hasta la fecha de corte para GL1 y GL3
no_ejecutados_gl1_new = contar_casos_no_ejecutados(plan_pruebas_integrales_df_new, 1)
no_ejecutados_gl3_new = contar_casos_no_ejecutados(plan_pruebas_integrales_df_new, 3)

# Calcular los casos ejecutados como la diferencia entre el total y los no ejecutados
ejecutados_gl1_new = unique_cases_gl1_new - no_ejecutados_gl1_new
ejecutados_gl3_new = unique_cases_gl3_new - no_ejecutados_gl3_new

# Contar los casos aprobados hasta la fecha de corte para GL1 y GL3
aprobados_gl1_new = contar_casos_aprobados(plan_pruebas_integrales_df_new, 1)
aprobados_gl3_new = contar_casos_aprobados(plan_pruebas_integrales_df_new, 3)

# Calcular el porcentaje de casos ejecutados hasta la fecha de corte
porcentaje_ejecutados_gl1 = calcular_porcentaje_ejecutados(unique_cases_gl1_new, ejecutados_gl1_new)
porcentaje_ejecutados_gl3 = calcular_porcentaje_ejecutados(unique_cases_gl3_new, ejecutados_gl3_new)

# Calcular el porcentaje planificado de forma dinámica
porcentaje_planificado_gl1 = calcular_porcentaje_planificado(plan_pruebas_integrales_df_new, 1)
porcentaje_planificado_gl3 = calcular_porcentaje_planificado(plan_pruebas_integrales_df_new, 3)

# Calcular la diferencia dinámica entre ejecutado y planificado
diferencia_gl1 = calcular_diferencia(porcentaje_ejecutados_gl1, porcentaje_planificado_gl1)
diferencia_gl3 = calcular_diferencia(porcentaje_ejecutados_gl3, porcentaje_planificado_gl3)

# Crear una tabla con los resultados incluyendo todos los campos anteriores
tabla_resultados_con_diferencia = pd.DataFrame({
    'N° GL': ['GL1', 'GL3'],
    'Total': [unique_cases_gl1_new, unique_cases_gl3_new],
    'Ejecutados': [ejecutados_gl1_new, ejecutados_gl3_new],
    'Pendientes (No ejecutados)': [no_ejecutados_gl1_new, no_ejecutados_gl3_new],
    '% Casos Ejecutados': [porcentaje_ejecutados_gl1, porcentaje_ejecutados_gl3],
    '% Casos Ejecutados Planificado': [porcentaje_planificado_gl1, porcentaje_planificado_gl3],
    'Diferencia': [diferencia_gl1, diferencia_gl3],
    'Avance Real (Aprobados)': [aprobados_gl1_new, aprobados_gl3_new]
}).applymap(centrar_contenido)

# Agregar la fila "Total"
totales = pd.DataFrame({
    'N° GL': ['Total'],
    'Total': [unique_cases_gl1_new + unique_cases_gl3_new],
    'Ejecutados': [ejecutados_gl1_new + ejecutados_gl3_new],
    'Pendientes (No ejecutados)': [no_ejecutados_gl1_new + no_ejecutados_gl3_new],
    '% Casos Ejecutados': [''],
    '% Casos Ejecutados Planificado': [''],
    'Diferencia': [''],
    'Avance Real (Aprobados)': [aprobados_gl1_new + aprobados_gl3_new]
}).applymap(centrar_contenido)

# Concatenar la fila de totales a la tabla
tabla_resultados_con_diferencia = pd.concat([tabla_resultados_con_diferencia, totales], ignore_index=True)

# Mostrar la tabla sin mostrar el índice
print(tabulate(tabla_resultados_con_diferencia, headers='keys', tablefmt='grid', showindex=False))
