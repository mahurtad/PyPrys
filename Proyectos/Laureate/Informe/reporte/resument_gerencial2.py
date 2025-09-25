import pandas as pd
from tabulate import tabulate

# Ruta del archivo actualizado
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar el archivo Excel y la hoja correspondiente
plan_pruebas_integrales_df_new = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Convertir la columna Fecha Inicial y Fecha Planificada a formato de fecha, ignorando errores
plan_pruebas_integrales_df_new['Fecha Inicial'] = pd.to_datetime(plan_pruebas_integrales_df_new['Fecha Inicial'], format='%d/%m/%Y', errors='coerce')
plan_pruebas_integrales_df_new['Fecha Planificada'] = pd.to_datetime(plan_pruebas_integrales_df_new['Fecha Planificada'], format='%d/%m/%Y', errors='coerce')

# Función para contar los casos no cancelados y en un rango de fechas, por GL
def calcular_avance_planificado(df, gl_value, fecha_inicio, fecha_fin):
    df_filtrado = df[
        (df['N° GL'] == gl_value) & 
        (df['Estado Entregable'] != 'Cancelado') &
        (df['Fecha Inicial'] >= fecha_inicio) & 
        (df['Fecha Inicial'] <= fecha_fin)
    ]
    casos_planificados = df_filtrado['N° CP'].nunique()
    return casos_planificados

# Definir la fecha más antigua en el campo "Fecha Inicial" como fecha de inicio
fecha_inicio_antigua = plan_pruebas_integrales_df_new['Fecha Inicial'].min()

# Definir la fecha de fin como la fecha de corte actual
fecha_fin_actual = pd.to_datetime('23/10/2024', format='%d/%m/%Y')

# Fecha de corte para Real Sem. Ant. (7 días antes)
fecha_corte_sem_ant = fecha_fin_actual - pd.Timedelta(days=7)

# Calcular el avance planificado desde las fechas más antiguas
avance_planificado_gl1_antigua = calcular_avance_planificado(plan_pruebas_integrales_df_new, 1, fecha_inicio_antigua, fecha_fin_actual)
avance_planificado_gl3_antigua = calcular_avance_planificado(plan_pruebas_integrales_df_new, 3, fecha_inicio_antigua, fecha_fin_actual)

# Recalcular los ejecutados hasta la fecha de corte
def contar_casos_ejecutados(df, gl_value, fecha_limite):
    df_filtrado = df[(df['N° GL'] == gl_value) & (df['Fecha Planificada'] <= fecha_limite)]
    casos_ejecutados = df_filtrado.groupby('N° CP').filter(
        lambda x: ((x['Estado Entregable'] == 'Aprobado') | (x['Estado Entregable'] == 'Pendiente retest')).any()
    )['N° CP'].nunique()
    return casos_ejecutados

# Calcular los casos ejecutados hasta la fecha de corte actual
ejecutados_gl1_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 1, fecha_fin_actual)
ejecutados_gl3_new = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 3, fecha_fin_actual)

# Calcular los ejecutados una semana antes (Real Sem. Ant.)
ejecutados_gl1_sem_ant = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 1, fecha_corte_sem_ant)
ejecutados_gl3_sem_ant = contar_casos_ejecutados(plan_pruebas_integrales_df_new, 3, fecha_corte_sem_ant)

# Real vs Planificado (Var) es la diferencia entre Avance Real y Avance Planificado
var_real_vs_planificado_gl1_antigua = ejecutados_gl1_new - avance_planificado_gl1_antigua
var_real_vs_planificado_gl3_antigua = ejecutados_gl3_new - avance_planificado_gl3_antigua

# Crear la tabla de resultados con el cálculo actualizado de Avance Planificado
tabla_resultados_actualizada = pd.DataFrame({
    'N° GL': ['GL1', 'GL3'],
    'Cant. Pruebas': [plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 1]['N° CP'].nunique(), 
                      plan_pruebas_integrales_df_new[plan_pruebas_integrales_df_new['N° GL'] == 3]['N° CP'].nunique()],
    'Avance Planificado': [avance_planificado_gl1_antigua, avance_planificado_gl3_antigua],
    f'Avance Real {fecha_fin_actual.strftime("%d/%m")}': [ejecutados_gl1_new, ejecutados_gl3_new],
    f'Real Sem. Ant. {fecha_corte_sem_ant.strftime("%d/%m")}': [ejecutados_gl1_sem_ant, ejecutados_gl3_sem_ant],
    'Real vs Planificado (Var)': [var_real_vs_planificado_gl1_antigua, var_real_vs_planificado_gl3_antigua]
})

# Mostrar la tabla sin mostrar el índice
print(tabulate(tabla_resultados_actualizada, headers='keys', tablefmt='grid', showindex=False))
