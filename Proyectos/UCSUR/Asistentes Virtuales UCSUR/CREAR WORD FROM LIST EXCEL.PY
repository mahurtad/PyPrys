import os
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT

# Rutas
carpeta_excel = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\Calendario Silabo\Documentos Word"
carpeta_word = os.path.join(carpeta_excel, "Documentos Word")
os.makedirs(carpeta_word, exist_ok=True)

# Función para agregar una tabla al documento Word desde un DataFrame
def agregar_tabla(doc, df):
    tabla = doc.add_table(rows=1, cols=len(df.columns))
    tabla.style = 'Table Grid'
    hdr_cells = tabla.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)

    for _, row in df.iterrows():
        row_cells = tabla.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)

# Recorre archivos Excel
for archivo in os.listdir(carpeta_excel):
    if archivo.lower().endswith(('.xlsx', '.xls')):
        ruta_excel = os.path.join(carpeta_excel, archivo)
        nombre_base = os.path.splitext(archivo)[0]
        ruta_word = os.path.join(carpeta_word, f"{nombre_base}.docx")

        # Crea documento Word vacío y horizontal
        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width

        # Lee solo la primera hoja del Excel
        try:
            df = pd.read_excel(ruta_excel)
            if not df.empty:
                agregar_tabla(doc, df)
        except Exception as e:
            doc.add_paragraph(f"No se pudo leer el archivo Excel: {str(e)}")

        # Guarda el Word
        doc.save(ruta_word)

print("✅ Word generado por cada Excel con una sola tabla (horizontal).")
