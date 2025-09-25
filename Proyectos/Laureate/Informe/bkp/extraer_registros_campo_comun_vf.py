import pandas as pd

# Definir las rutas de los archivos
file_path = r'C:\Users\manue\Downloads\reporte_pendientes.xlsx'
output_file_path = r'C:\Users\manue\Downloads\reporte_cruce3.xlsx'

# Cargar las hojas necesarias del archivo
lista_total_casos = pd.read_excel(file_path, sheet_name='Lista Total CP-Procesos')
pendiente_gl1 = pd.read_excel(file_path, sheet_name='Pendiente')

# Nombre de la columna para hacer el cruce
case_column = 'N° del Caso de Prueba'

# Convertir los valores de "N° del Caso de Prueba" a cadenas de texto para mantener el formato exacto
lista_total_casos[case_column] = lista_total_casos[case_column].astype(str)
pendiente_gl1[case_column] = pendiente_gl1[case_column].astype(str)

# Filtrar los registros de "Lista Total Casos" que coinciden con los de "Pendiente"
filtered_gl1 = lista_total_casos[lista_total_casos[case_column].isin(pendiente_gl1[case_column])]

# Exportar el resultado en un archivo Excel con la hoja "PendienteGL1"
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    filtered_gl1.to_excel(writer, sheet_name='PendienteGL1', index=False)

print(f"El reporte ha sido exportado exitosamente a: {output_file_path}")
