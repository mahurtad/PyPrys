import pandas as pd

# Ruta del archivo original
file_path = r"C:\Users\manue\OneDrive - EduCorpPERU\UPC\GEMA (BANNER)\Documentos\Bitacora de Observaciones GEMA.xlsx"

# Campos a extraer
fields_to_extract = [
    "N° Caso de Prueba",
    "Descripción de Observación",
    "Tipo de Observación",
    "Fecha Ejecución",
    "Estado",
    "VB Usuario",
    "Recategorización"
]

# Nombre de la hoja
sheet_name = "QA"

# Ruta del archivo de salida
output_path = r"C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Reporte_Bitacora.xlsx"

# Proceso de extracción y exportación
try:
    # Leer los datos desde la hoja especificada
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Extraer los campos requeridos
    extracted_data = data[fields_to_extract]
    
    # Contar observaciones por tipo
    type_counts = extracted_data["Tipo de Observación"].value_counts().reset_index()
    type_counts.columns = ["Tipo de Observación", "Cantidad"]
    
    # Exportar a un nuevo archivo Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Detalle")
        type_counts.to_excel(writer, index=False, sheet_name="Resumen")
    
    print(f"Reporte generado exitosamente: {output_path}")
except Exception as e:
    print(f"Error: {e}")
