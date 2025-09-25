import os
import pandas as pd
from docx import Document

def extraer_descripcion_de_tablas(doc):
    """
    Busca la celda 'Descripción de la actividad' dentro de todas las tablas del documento
    y devuelve el contenido de la celda adyacente (derecha) si existe.
    """
    for table in doc.tables:
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                texto = cell.text.strip().lower()
                if "descripción de la actividad" in texto:
                    if i + 1 < len(row.cells):
                        contenido = row.cells[i + 1].text.strip()
                        if contenido:
                            return contenido
    return "Descripción no encontrada"

def procesar_documentos(ruta_carpeta):
    """
    Procesa todos los archivos .docx en la ruta dada y extrae las descripciones.
    """
    registros = []

    for archivo in os.listdir(ruta_carpeta):
        if archivo.lower().endswith(".docx"):
            ruta_completa = os.path.join(ruta_carpeta, archivo)
            try:
                doc = Document(ruta_completa)
                descripcion = extraer_descripcion_de_tablas(doc)
            except Exception as e:
                descripcion = f"Error al procesar: {e}"
            registros.append({
                "Evaluación": os.path.splitext(archivo)[0],
                "Descripción de la actividad": descripcion
            })

    return pd.DataFrame(registros)

def exportar_excel(df, ruta_archivo):
    """
    Exporta el DataFrame resultante a un archivo Excel.
    """
    df.to_excel(ruta_archivo, index=False)
    print(f"✅ Excel generado correctamente en:\n{ruta_archivo}")

# ========== CONFIGURACIÓN DEL USUARIO ==========
if __name__ == "__main__":
    carpeta_docx = r"C:\Users\manue\Downloads\Rúbricas oficiales\Turismo Sostenible y Hotelería - Administración hotelera y turismo - adm de empresas\Dirección Estratégica para el sector turístico y hotelero"
    salida_excel = os.path.join(carpeta_docx, "descripciones_actividades.xlsx")

    df = procesar_documentos(carpeta_docx)
    exportar_excel(df, salida_excel)
