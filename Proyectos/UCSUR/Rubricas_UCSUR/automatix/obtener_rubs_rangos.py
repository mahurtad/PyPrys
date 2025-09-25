import os
import re
from docx import Document
import pandas as pd

# Ruta base donde se encuentran los documentos de Word
base_folder_path = r"C:\Users\manue\Downloads\Agricultura Ecológica"

# Expresión regular para identificar rangos de calificación
RANGO_REGEX = r"(\d+)\s*-\s*(\d+)"

# Lista para almacenar los resultados
rangos_encontrados = []

def buscar_rangos_en_tabla(table):
    """Busca rangos de calificación en una tabla de Word."""
    rangos = []
    for row in table.rows:
        for cell in row.cells:
            texto = cell.text.strip()
            matches = re.findall(RANGO_REGEX, texto)
            if matches:
                for match in matches:
                    rangos.append((int(match[0]), int(match[1])))
    return rangos

def procesar_documento(word_file_path):
    """Procesa un documento de Word para buscar rangos de calificación en sus tablas."""
    doc = Document(word_file_path)
    tables = doc.tables
    resultados = []

    if not tables:
        return None

    for table_index, table in enumerate(tables):
        rangos = buscar_rangos_en_tabla(table)
        if rangos:
            resultados.append({
                "Archivo": word_file_path,
                "Tabla": table_index + 1,
                "Rangos": rangos
            })

    return resultados

# Recorrer todas las carpetas y subcarpetas
for root, dirs, files in os.walk(base_folder_path):
    word_files = [f for f in files if f.endswith(".docx")]
    
    for filename in word_files:
        word_file_path = os.path.join(root, filename)
        resultados = procesar_documento(word_file_path)
        
        if resultados:
            rangos_encontrados.extend(resultados)

# Generar un reporte en Excel con los rangos encontrados
if rangos_encontrados:
    df_report = pd.DataFrame(rangos_encontrados)
    report_path = os.path.join(base_folder_path, "Reporte_Rangos_Calificacion.xlsx")
    df_report.to_excel(report_path, index=False)
    print(f"✅ Reporte generado en: {report_path}")
else:
    print("✅ No se encontraron rangos de calificación en los documentos procesados.")