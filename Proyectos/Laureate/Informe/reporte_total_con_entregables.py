import pandas as pd
import os

# Ruta del archivo Excel original
file_path_new = r'C:\Users\manue\OneDrive - EduCorpPERU\Calidad de Software - UPC\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'


# Cargar el archivo Excel y la hoja correspondiente
df = pd.read_excel(file_path_new, sheet_name='Plan Pruebas Integrales')

# Asegurarse de que los nombres de las columnas estén limpios (sin espacios extra)
df.columns = df.columns.str.strip()

# Convertir las columnas de fechas al formato de fecha, ignorando errores
df['Fecha Planificada'] = pd.to_datetime(df['Fecha Planificada'], errors='coerce')

# Definir la función para asignar el "Estado de CP"
def asignar_estado(row):
    if row['Estado Entregable'] == 'Aprobado':
        return 'Aprobado'
    elif row['Estado Entregable'] == 'No ejecutado':
        return 'No ejecutado'
    else:
        return 'Pendiente de Retest'

# Aplicar la función de estado
df['Estado de CP'] = df.apply(asignar_estado, axis=1)

# Filtrar y seleccionar únicamente las columnas necesarias
columnas_seleccionadas = ['N° GL', 'N° CP', 'N° D', 'Caso de Prueba', 'Nombre de Tarea (Desarrollo)',
                          'Fecha Planificada', 'BPO', 'PM Asignado', 'Periférico', 'Estado de CP']
df_seleccion = df[columnas_seleccionadas]

# Exportar los resultados a un archivo Excel
output_path = os.path.join(os.path.dirname(file_path_new), 'plan_pruebas_integrales_250222025.xlsx')
df_seleccion.to_excel(output_path, index=False)

print(f"El archivo con los registros seleccionados ha sido guardado en: {output_path}")
