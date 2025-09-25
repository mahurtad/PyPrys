import pandas as pd

# Ruta del archivo original
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar el archivo original con pandas
plan_pruebas_integrales_df_new = pd.read_excel(file_path_new, sheet_name=None)  # Carga todas las hojas existentes

# Filtrar los casos aprobados para GL1 y GL3
def obtener_casos_aprobados_por_gl(df, gl_value):
    # Filtrar por GL y casos donde todos los registros del caso de prueba están aprobados
    df_filtrado = df[df['N° GL'] == gl_value]
    casos_aprobados = df_filtrado.groupby('N° CP').filter(
        lambda x: (x['Estado Entregable'] == 'Aprobado').all()
    )
    return casos_aprobados

# Obtener los casos aprobados para GL1 y GL3
plan_pruebas_integrales_df = plan_pruebas_integrales_df_new['Plan Pruebas Integrales']
casos_aprobados_gl1 = obtener_casos_aprobados_por_gl(plan_pruebas_integrales_df, 1)
casos_aprobados_gl3 = obtener_casos_aprobados_por_gl(plan_pruebas_integrales_df, 3)

# Agregar las nuevas hojas de casos aprobados al DataFrame existente
plan_pruebas_integrales_df_new['Casos Aprobados GL1'] = casos_aprobados_gl1
plan_pruebas_integrales_df_new['Casos Aprobados GL3'] = casos_aprobados_gl3

# Guardar todas las hojas en un nuevo archivo Excel
output_file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3_actualizado.xlsx'
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for sheet_name, df_sheet in plan_pruebas_integrales_df_new.items():
        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"El archivo ha sido guardado exitosamente en: {output_file_path}")
