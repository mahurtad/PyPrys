import pandas as pd

# Ruta del archivo de entrada
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'

# Cargar la hoja relevante
plan_pruebas_data = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales')

# Asegurar que la columna de fecha esté en formato datetime
plan_pruebas_data['Fecha Planificada'] = pd.to_datetime(plan_pruebas_data['Fecha Planificada'], errors='coerce')

# Definir los rangos de fechas
date_ranges = [
    ('2024-11-09', '2024-11-15'),
    ('2024-11-16', '2024-11-22'),
    ('2024-11-23', '2024-11-29'),
    ('2024-11-30', '2024-12-06'),
    ('2024-12-07', '2024-12-13'),
    ('2024-12-14', '2024-12-20'),
    ('2024-12-21', '2024-12-27'),
    ('2024-12-28', '2025-01-05'),
    ('2025-01-06', '2025-01-12'),
    ('2025-01-13', '2025-01-19'),
    ('2025-01-20', '2025-01-26'),
    ('2025-01-27', '2025-02-02'),
    ('2025-02-03', '2025-02-09'),
    ('2025-02-10', '2025-02-16'),
    ('2025-02-17', '2025-02-24'),
    ('2025-02-25', '2025-03-02')
]

# Convertir los rangos de fechas a un DataFrame para facilitar el manejo
date_ranges_df = pd.DataFrame(date_ranges, columns=['Fecha Inicio', 'Fecha Fin'])
date_ranges_df['Fecha Inicio'] = pd.to_datetime(date_ranges_df['Fecha Inicio'])
date_ranges_df['Fecha Fin'] = pd.to_datetime(date_ranges_df['Fecha Fin'])

# Función para calcular los casos únicos y cumplir las reglas definidas
def recompute_projections_unique_cases(data, gl):
    results = []
    seen_cases = set()  # Casos ya considerados en rangos anteriores
    for _, row in date_ranges_df.iterrows():
        # Filtrar datos por rango de fechas y GL
        filtered_data = data[
            (data['Fecha Planificada'] >= row['Fecha Inicio']) & 
            (data['Fecha Planificada'] <= row['Fecha Fin']) & 
            (data['N° GL'] == gl)
        ]
        # Identificar casos donde todos los registros sean "No ejecutado"
        valid_test_cases = filtered_data.groupby('N° CP').filter(
            lambda group: (group['Estado Entregable'] == 'No ejecutado').all()
        )
        # Considerar solo la fecha más reciente para cada caso
        latest_test_cases = valid_test_cases.loc[
            valid_test_cases.groupby('N° CP')['Fecha Planificada'].idxmax()
        ]
        # Excluir casos ya considerados en rangos anteriores
        new_test_cases = latest_test_cases[~latest_test_cases['N° CP'].isin(seen_cases)]
        # Actualizar los casos ya vistos
        seen_cases.update(new_test_cases['N° CP'].dropna().unique())
        # Filtrar casos que realmente caen en el rango actual
        final_test_cases = new_test_cases[
            (new_test_cases['Fecha Planificada'] >= row['Fecha Inicio']) & 
            (new_test_cases['Fecha Planificada'] <= row['Fecha Fin'])
        ]
        # Construir el resultado
        test_cases = ', '.join(final_test_cases['N° CP'].dropna().astype(str).unique())
        case_count = len(final_test_cases['N° CP'].dropna().unique())
        results.append({
            'Fecha Inicio': row['Fecha Inicio'],
            'Fecha Fin': row['Fecha Fin'],
            'Cantidad de Casos': case_count,
            'Casos de Prueba': test_cases if test_cases else ''
        })
    return pd.DataFrame(results)

# Calcular proyecciones para GL1 y GL3
gl1_projection_unique = recompute_projections_unique_cases(plan_pruebas_data, 1)
gl1_projection_unique['Go Live'] = 'GL1'

gl3_projection_unique = recompute_projections_unique_cases(plan_pruebas_data, 3)
gl3_projection_unique['Go Live'] = 'GL3'

# Dar formato a las fechas como DD/MM/AAAA
gl1_projection_unique['Fecha Inicio'] = gl1_projection_unique['Fecha Inicio'].dt.strftime('%d/%m/%Y')
gl1_projection_unique['Fecha Fin'] = gl1_projection_unique['Fecha Fin'].dt.strftime('%d/%m/%Y')
gl3_projection_unique['Fecha Inicio'] = gl3_projection_unique['Fecha Inicio'].dt.strftime('%d/%m/%Y')
gl3_projection_unique['Fecha Fin'] = gl3_projection_unique['Fecha Fin'].dt.strftime('%d/%m/%Y')

# Ruta para guardar el archivo de salida
output_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\proyeccion_casos_28112024.xlsx'

# Exportar a Excel
with pd.ExcelWriter(output_path) as writer:
    gl1_projection_unique.to_excel(writer, sheet_name='GL1_Projection', index=False)
    gl3_projection_unique.to_excel(writer, sheet_name='GL3_Projection', index=False)

print(f"Proyecciones exportadas exitosamente a: {output_path}")
