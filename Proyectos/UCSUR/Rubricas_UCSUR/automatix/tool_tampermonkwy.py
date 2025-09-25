import os
import pandas as pd
from docx import Document

def extract_tables_from_word(doc_path):
    """Extrae todas las tablas de un documento de Word y las guarda en un DataFrame de Pandas."""
    doc = Document(doc_path)
    tables = []
    
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        df = pd.DataFrame(data)
        tables.append(df)
    
    return tables

def save_tables_to_excel(tables, excel_path):
    """Guarda las tablas extraídas en un archivo de Excel, cada una en una hoja diferente."""
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, df in enumerate(tables):
                sheet_name = f'Tabla_{i+1}'
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        print(f"Tablas guardadas en: {excel_path}")
    except PermissionError:
        print("Error: No se pudo escribir en el archivo. Asegúrate de que no está abierto y que tienes permisos.")
    except Exception as e:
        print(f"Error inesperado: {e}")

def main():
    # Ruta del documento Word
    word_path = r"C:\\Users\\manue\\OneDrive - Grupo Educad\\3. Documentos validados\\2. Enfermería\\El Cuidado de Enfermería\\EC1_El Cuidado de Enfermería_Enfermería.docx"
    
    # Verificar si el archivo de Word existe
    if not os.path.exists(word_path):
        print(f"Error: El archivo no se encuentra en la ruta especificada: {word_path}")
        return
    
    # Obtener la ruta base y definir el nombre del archivo Excel
    base_dir = os.path.expanduser("~\\Documents")  # Guardar en Documentos para evitar problemas de permisos
    excel_path = os.path.join(base_dir, "EC1_El_Cuidado.xlsx")
    
    # Extraer tablas y guardarlas en Excel
    tables = extract_tables_from_word(word_path)
    if tables:
        save_tables_to_excel(tables, excel_path)
    else:
        print("No se encontraron tablas en el documento.")

if __name__ == "__main__":
    main()