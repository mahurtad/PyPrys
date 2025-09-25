import pandas as pd

# Definir la ruta del archivo y el nombre de la hoja
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
sheet_name = 'Plan Pruebas Integrales'
column_name = 'N° GL'

# Leer el archivo Excel y omitir la primera fila si es necesario
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Contar los registros no nulos en la columna especificada, omitiendo la cabecera
count_records = df[column_name].iloc[1:].count()  # Iloc[1:] excluye la primera fila (cabecera)

print(f'Cantidad de registros en la columna {column_name}: {count_records}')

