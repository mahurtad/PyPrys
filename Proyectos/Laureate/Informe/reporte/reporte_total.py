import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Ruta del archivo en tu sistema local
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
sheet_name = 'Control General GL1-GL3'

# Cargar los datos desde la hoja especificada
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Mostrar una vista general de los datos cargados (primeras filas)
print(data.head())

# Generar un gráfico de barras para la cantidad de casos por estado (asumiendo que tienes una columna "Estado")
plt.figure(figsize=(10, 6))
sns.countplot(x='Estado Caso de Prueba', data=data)
plt.title('Cantidad de Casos por Estado')
plt.xticks(rotation=45)
plt.tight_layout()

# Guardar el gráfico como archivo
plt.savefig(r'C:\Users\manue\Downloads\grafico_casos_por_estado.png')

# Mostrar el gráfico
plt.show()

# Crear una tabla resumen de la cantidad de casos por estado
tabla_resumen = data['Estado'].value_counts().reset_index()
tabla_resumen.columns = ['Estado', 'Cantidad de Casos']

# Mostrar la tabla de resumen
print(tabla_resumen)

# Exportar la tabla de resumen a un archivo Excel
tabla_resumen.to_excel(r'C:\Users\manue\Downloads\reporte_tabla_resumen_.xlsx', index=False)

print(f"El reporte ha sido generado y guardado en: C:\\Users\\manue\\Downloads")
