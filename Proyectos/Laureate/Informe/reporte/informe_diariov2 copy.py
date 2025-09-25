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

# Función para contar los casos aprobados hasta la fecha de corte (un caso es aprobado solo si **todos** los registros tienen estado Aprobado)
def contar_casos_aprobados(df, gl_value, fecha_limite):
    df_filtrado = df[(df['N° GL'] == gl_value) & (df['Fecha Planificada'] <= fecha_limite)]
    casos_aprobados = df_filtrado.groupby('N° CP').filter(
        lambda x: all(x['Estado Entregable'] == 'Aprobado')  # Todos los registros deben estar aprobados
    )['N° CP'].nunique()
    return casos_aprobados

# Función para contar los casos ejecutados hasta la fecha de corte (casos con al menos un registro aprobado o en retest)
def contar_casos_ejecutados(df, gl_value, fecha_limite):
    df_filtrado = df[(df['N° GL'] == gl_value) & (df['Fecha Planificada'] <= fecha_limite)]
    casos_ejecutados = df_filtrado.groupby('N° CP').filter(
        lambda x: ((x['Estado Entregable'] == 'Aprobado') | (x['Estado Entregable'] == 'Pendiente retest')).any()
    )['N° CP'].nunique()
    return casos_ejecutados

# Función para contar los casos pendientes (Total - Ejecutados)
def contar_casos_pendientes(total_cases, ejecutados_cases):
    return total_cases - ejecutados_cases

# Función para calcular el porcentaje de casos ejecutados (Ejecutados / Total) * 100 y mostrarlo como porcentaje
def calcular_porcentaje_ejecutados(total_cases, ejecutados_cases):
    return "{:.2f}%".format((ejecutados_cases / total_cases) * 100) if total_cases > 0 else "0%"

# Función para contar los casos ejecutados planificados hasta la fecha de corte
def contar_casos_ejecutados_planificado(df, gl_value, fecha_limite):
    df_filtrado = df[(df['N° GL'] == gl_value) & (df['Fecha Planificada'] <= fecha_limite)]
    casos_ejecutados_planificados = df_filtrado.groupby('N° CP').filter(
        lambda x: ((x['Estado Entregable'] == 'Aprobado') | (x['Estado Entregable'] == 'Pendiente retest')).any()
    )['N° CP'].nunique()
    return casos_ejecutados_planificados

# Fecha límite para el cálculo (esta será la fecha de corte)
fecha_limite = pd.to_datetime('23/10/2024', format='%d/%m/%Y')

# Contar los casos de prueba únicos (N° CP) para GL1 y GL3 (el Total no varía)
unique_cases_gl1_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 1]['N° CP'].nunique()
unique_cases_gl3_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 3]['N° CP'].nunique()

# Contar los casos aprobados hasta la fecha de corte para GL1 y GL3 (con la nueva lógica de aprobación total)
approved_cases_gl1_new = contar_casos_aprobados(plan_pruebas_integrales_df_new, 1, fecha_limite)
approved_cases_gl3_new = contar_casos_aprobados(plan_pruebas_integrales_df_new, 3, fecha_limite)

# Contar los casos ejecutados hasta la fecha de corte para GL1 y GL3
ejecutados_gl1_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 1, fecha_limite)
ejecutados_gl3_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 3, fecha_limite)

# Calcular los casos pendientes (Total - Ejecutados)
pendientes_gl1_new = contar_casos_pendientes(unique_cases_gl1_new, ejecutados_gl1_new)
pendientes_gl3_new = contar_casos_pendientes(unique_cases_gl3_new, ejecutados_gl3_new)

# Calcular el porcentaje de casos ejecutados hasta la fecha de corte
porcentaje_ejecutados_gl1 = calcular_porcentaje_ejecutados(unique_cases_gl1_new, ejecutados_gl1_new)
porcentaje_ejecutados_gl3 = calcular_porcentaje_ejecutados(unique_cases_gl3_new, ejecutados_gl3_new)

# Calcular el porcentaje de casos ejecutados planificados hasta la fecha de corte
ejecutados_planificados_gl1 = contar_casos_ejecutados_planificado(plan_pruebas_integrales_df_new, 1, fecha_limite)
ejecutados_planificados_gl3 = contar_casos_ejecutados_planificado(plan_pruebas_integrales_df_new, 3, fecha_limite)
porcentaje_ejecutados_planificado_gl1 = calcular_porcentaje_ejecutados(unique_cases_gl1_new, ejecutados_planificados_gl1)
porcentaje_ejecutados_planificado_gl3 = calcular_porcentaje_ejecutados(unique_cases_gl3_new, ejecutados_planificados_gl3)

# Función para calcular la diferencia entre % Casos Ejecutados y % Casos Ejecutados Planificado
def calcular_diferencia(ejecutados, planificado):
    return "{:.2f}%".format(float(ejecutados.strip('%')) - float(planificado.strip('%')))

# Calcular la diferencia para GL1 y GL3
diferencia_gl1 = calcular_diferencia(porcentaje_ejecutados_gl1, porcentaje_ejecutados_planificado_gl1)
diferencia_gl3 = calcular_diferencia(porcentaje_ejecutados_gl3, porcentaje_ejecutados_planificado_gl3)

# Crear una tabla con los resultados, renombrando la columna "Aprobados" a "Avance Real (Aprobados)"
tabla_resultados_con_diferencia = pd.DataFrame({
    'N° GL': ['GL1', 'GL3'],
    'Total': [unique_cases_gl1_new, unique_cases_gl3_new],
    'Ejecutados': [ejecutados_gl1_new, ejecutados_gl3_new],
    'Pendientes': [pendientes_gl1_new, pendientes_gl3_new],
    '% Casos Ejecutados': [porcentaje_ejecutados_gl1, porcentaje_ejecutados_gl3],
    '% Casos Ejecutados Planificado': [porcentaje_ejecutados_planificado_gl1, porcentaje_ejecutados_planificado_gl3],
    'Diferencia': [diferencia_gl1, diferencia_gl3],
    'Avance Real (Aprobados)': [approved_cases_gl1_new, approved_cases_gl3_new]
}).applymap(centrar_contenido)

# Mostrar la tabla sin mostrar el índice
print(tabulate(tabla_resultados_con_diferencia, headers='keys', tablefmt='grid', showindex=False))
