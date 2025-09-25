import pandas as pd
from tabulate import tabulate

# Definir la ruta del archivo y el nombre de la hoja
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
sheet_name = 'Plan Pruebas Integrales'
column_name = 'Fecha Replanificada CP' 
# Leer el archivo Excel
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Asegurarnos de que la columna de fecha se convierta al formato de fecha de pandas
df[column_name] = pd.to_datetime(df[column_name], errors='coerce', dayfirst=True)

# Solicitar al usuario la fecha que desea filtrar
filter_date = input("Por favor, ingrese la fecha que desea consultar (formato DD/MM/AAAA): ")

# Convertir la fecha ingresada a un objeto de fecha de pandas
try:
    filter_date = pd.to_datetime(filter_date, format='%d/%m/%Y')
except ValueError:
    print("Formato de fecha incorrecto. Por favor ingrese en formato DD/MM/AAAA.")
    exit()

# Filtrar las filas donde la columna 'Fecha Caso Uso' coincida con la fecha ingresada
filtered_rows = df[df[column_name] == filter_date]

# Verificar si se encontraron filas
if not filtered_rows.empty:
    # Seleccionar solo las columnas específicas que deseas mostrar
    columns_to_display = ['N° GL', 'N° CP', 'N° D', 'Fecha Replanificada CP', 'BPO', 'PM Asignado', 'Periférico', 'Estado CP', 'Hora de Inicio', 'Hora de Fin']
    filtered_rows = filtered_rows[columns_to_display]
    
    # Mostrar todas las columnas en pandas
    pd.set_option('display.max_columns', None)
    
    # Imprimir las filas filtradas con formato de tabla usando `tabulate`
    print(f"\nFilas que contienen la fecha {filter_date.strftime('%d/%m/%Y')} según 'Fecha Caso de Prueba':\n")
    print(tabulate(filtered_rows, headers='keys', tablefmt='grid', showindex=False))
else:
    print(f"No se encontraron filas con la fecha {filter_date.strftime('%d/%m/%Y')} en 'Fecha Caso de Prueba'.")
