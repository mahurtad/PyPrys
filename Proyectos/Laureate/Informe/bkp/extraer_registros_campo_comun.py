import pandas as pd

# Definir las rutas de los archivos
file_path = r'C:\Users\manue\Downloads\no_ejecutados.xlsx'
output_file_path = r'C:\Users\manue\Downloads\reporte_pendientes.xlsx'

# Cargar las hojas necesarias del archivo
lista_total_casos = pd.read_excel(file_path, sheet_name='Lista Total Casos')
pendiente_gl1 = pd.read_excel(file_path, sheet_name='PendienteGL1')
pendiente_gl3 = pd.read_excel(file_path, sheet_name='PendienteGL3')

# Nombre de la columna para hacer el cruce
case_column = 'N° del Caso de Prueba'

# Filtrar los registros de "Lista Total Casos" que coinciden con los de "PendienteGL1"
filtered_gl1 = lista_total_casos[lista_total_casos[case_column].isin(pendiente_gl1[case_column])]

# Filtrar los registros de "Lista Total Casos" que coinciden con los de "PendienteGL3"
filtered_gl3 = lista_total_casos[lista_total_casos[case_column].isin(pendiente_gl3[case_column])]

# Exportar el resultado en un archivo Excel con dos hojas "PendienteGL1" y "PendienteGL3"
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    filtered_gl1.to_excel(writer, sheet_name='PendienteGL1', index=False)
    filtered_gl3.to_excel(writer, sheet_name='PendienteGL3', index=False)

print(f"El reporte ha sido exportado exitosamente a: {output_file_path}")
