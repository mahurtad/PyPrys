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

# Función para contar los casos no ejecutados
def contar_casos_no_ejecutados(df, gl_value):
    df_filtrado = df[df['N° GL'] == gl_value]
    casos_no_ejecutados = df_filtrado.groupby('N° CP').filter(
        lambda x: (x['Estado Entregable'] == 'No ejecutado').all()
    )['N° CP'].nunique()
    return casos_no_ejecutados

# Función para contar los casos ejecutados (al menos un registro tiene un estado diferente de "No Ejecutado")
def contar_casos_ejecutados(df, gl_value):
    df_filtrado = df[df['N° GL'] == gl_value]
    casos_ejecutados = df_filtrado.groupby('N° CP').filter(
        lambda x: not all(x['Estado Entregable'] == 'No ejecutado')
    )['N° CP'].nunique()
    return casos_ejecutados

# Función para contar los casos pendientes (Total - Ejecutados)
def contar_casos_pendientes(total_cases, ejecutados_cases):
    return total_cases - ejecutados_cases

# Función para calcular el porcentaje de casos ejecutados (Ejecutados / Total) * 100
def calcular_porcentaje_ejecutados(total_cases, ejecutados_cases):
    return "{:.2f}%".format((ejecutados_cases / total_cases) * 100) if total_cases > 0 else "0%"

# Fecha límite para el cálculo (esta será la fecha de corte)
fecha_limite = pd.to_datetime('24/10/2024', format='%d/%m/%Y')

# Contar los casos de prueba únicos (N° CP) para GL1 y GL3
unique_cases_gl1_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 1]['N° CP'].nunique()
unique_cases_gl3_new = plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 3]['N° CP'].nunique()

# Contar los casos ejecutados hasta la fecha de corte para GL1 y GL3
ejecutados_gl1_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 1)
ejecutados_gl3_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 3)

# Calcular los casos pendientes (Total - Ejecutados)
pendientes_gl1_new = contar_casos_pendientes(unique_cases_gl1_new, ejecutados_gl1_new)
pendientes_gl3_new = contar_casos_pendientes(unique_cases_gl3_new, ejecutados_gl3_new)

# Calcular el porcentaje de casos ejecutados hasta la fecha de corte
porcentaje_ejecutados_gl1 = calcular_porcentaje_ejecutados(unique_cases_gl1_new, ejecutados_gl1_new)
porcentaje_ejecutados_gl3 = calcular_porcentaje_ejecutados(unique_cases_gl3_new, ejecutados_gl3_new)

# Crear una tabla con los resultados incluyendo todos los campos anteriores
tabla_resultados_con_diferencia = pd.DataFrame({
    'N° GL': ['GL1', 'GL3'],
    'Total': [unique_cases_gl1_new, unique_cases_gl3_new],
    'Ejecutados': [ejecutados_gl1_new, ejecutados_gl3_new],
    'Pendientes': [pendientes_gl1_new, pendientes_gl3_new],
    '% Casos Ejecutados': [porcentaje_ejecutados_gl1, porcentaje_ejecutados_gl3],
    '% Casos Ejecutados Planificado': ["69.47%", "14.29%"],  # Ejemplos
    'Diferencia': ["12.64%", "10.20%"],  # Ejemplos
    'Avance Real (Aprobados)': [65, 7]  # Ejemplos
}).applymap(centrar_contenido)

# Mostrar la tabla sin mostrar el índice
print(tabulate(tabla_resultados_con_diferencia, headers='keys', tablefmt='grid', showindex=False))
