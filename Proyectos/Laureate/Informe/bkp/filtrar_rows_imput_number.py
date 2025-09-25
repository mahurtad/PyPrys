import pandas as pd
from tabulate import tabulate

# Definir la ruta del archivo y el nombre de la hoja
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
sheet_name = 'Plan Pruebas Integrales'
column_name = 'N° CU'

# Leer el archivo Excel
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Solicitar al usuario el número que desea filtrar
filter_number = int(input("Por favor, ingrese el N° CU que desea consultar: "))

# Filtrar las filas donde la columna 'N° CU' tenga el valor específico
filtered_rows = df[df[column_name] == filter_number]

# Verificar si se encontraron filas
if not filtered_rows.empty:
    # Seleccionar solo las columnas específicas que deseas mostrar
    columns_to_display = ['N° GL', 'N° CU', 'N° D', 'Fecha Caso Uso', 'BPO', 'PM Asignado', 'Periférico', 'Estado CU', 'Hora de Inicio', 'Hora de Fin']
    filtered_rows = filtered_rows[columns_to_display]
    
    # Mostrar todas las columnas en pandas
    pd.set_option('display.max_columns', None)
    
    # Imprimir las filas filtradas con formato de tabla usando `tabulate`
    print(f"\nFilas que contienen el N° CU {filter_number}:\n")
    print(tabulate(filtered_rows, headers='keys', tablefmt='grid', showindex=False))
else:
    print(f"No se encontraron filas con el N° CU {filter_number}.")
