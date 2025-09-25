import os
import pandas as pd

# Ruta del archivo Excel y la carpeta de PDFs
ruta_carpeta_pdfs = r'C:\Users\manue\Downloads\Generador\TALLER TRANSFORMANDO LAS EVALUACIONES CON INTELIGENCIA ARTIFICIAL\Certificados'
ruta_excel = r'C:\Users\manue\Downloads\Generador\Nombres.xlsx'

# Leer el archivo Excel, específicamente la hoja "Actividad6"
df = pd.read_excel(ruta_excel, sheet_name='Actividad5')

# Asegurarse de que las columnas se llaman correctamente como en el enunciado
columna_pdf = 'PDF'      # Columna que contiene el número del PDF
columna_nombre = 'Nombre' # Columna que contiene el nombre con el que renombrar el archivo

# Iterar sobre las filas del DataFrame para renombrar los archivos
for index, row in df.iterrows():
    numero_pdf = str(row[columna_pdf])  # Asegurarse de que sea una cadena
    nuevo_nombre = row[columna_nombre]  # Obtener el nuevo nombre

    # Construir la ruta del archivo PDF actual
    archivo_pdf_actual = os.path.join(ruta_carpeta_pdfs, f'{numero_pdf}.pdf')
    
    # Construir la nueva ruta con el nuevo nombre
    nuevo_archivo_pdf = os.path.join(ruta_carpeta_pdfs, f'{nuevo_nombre}.pdf')

    # Verificar si el archivo existe y luego renombrarlo
    if os.path.exists(archivo_pdf_actual):
        os.rename(archivo_pdf_actual, nuevo_archivo_pdf)
        print(f'Renombrado: {archivo_pdf_actual} -> {nuevo_archivo_pdf}')
    else:
        print(f'El archivo {archivo_pdf_actual} no existe.')

print('Proceso completado.')
