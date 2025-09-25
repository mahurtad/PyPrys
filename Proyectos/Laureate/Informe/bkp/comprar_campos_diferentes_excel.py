import pandas as pd

# Define la ruta del archivo y el nombre de la hoja
file_path = r'C:\Users\manue\Downloads\Estados de CPs GL1-GL2 Banner.xlsx'
sheet_name = 'Conteo'

# Cargar los datos desde el archivo Excel
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Nombres de las columnas
executed_column = 'Ejecutados'
total_cases_column = 'Total de Casos'

# Convertir las columnas a listas, eliminando valores nulos
executed = data[executed_column].dropna().tolist()
total_cases = data[total_cases_column].dropna().tolist()

# Encontrar los valores que están en 'Total de Casos' pero no en 'Ejecutados'
not_executed = [case for case in total_cases if case not in executed]

# Crear un DataFrame con los resultados
result_df = pd.DataFrame(not_executed, columns=['No Ejecutados (En Total de Casos)'])

# Exportar el DataFrame resultante a un archivo Excel
output_file_path = r'C:\Users\manue\Downloads\no_ejecutados_gl3.xlsx'
result_df.to_excel(output_file_path, index=False)

print(f"El archivo ha sido exportado exitosamente a: {output_file_path}")
