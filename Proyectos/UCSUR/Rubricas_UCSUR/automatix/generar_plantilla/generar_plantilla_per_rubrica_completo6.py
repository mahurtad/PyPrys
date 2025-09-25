import os
import re
import pandas as pd
from docx import Document
from openpyxl import load_workbook

# === CONFIGURACIÓN ===
base_folder_path = r"C:\Users\manue\Downloads\Medicina Veterinaria y Zootecnia - Copy\Tecnología e Industrialización de Alimentos\New folder"
error_log = []
validation_log = []
document_log = []

# === MAPA DE EQUIVALENCIAS ===
mapa_equivalencias = {
    "EC1": "Evaluación Continua 1", "EC2": "Evaluación Continua 2", 
    "EC3": "Evaluación Continua 3", "EC4": "Evaluación Continua 4",
    "EF": "Evaluación Final", "EP": "Evaluación Parcial", 
    "ED": "Evaluación Diagnóstica",
    "CONTINUA1": "Evaluación Continua 1", "CONTINUA2": "Evaluación Continua 2", 
    "CONTINUA3": "Evaluación Continua 3", "CONTINUA4": "Evaluación Continua 4", 
    "DIAGNOSTICA": "Evaluación Diagnóstica", 
    "FINAL": "Evaluación Final", "PARCIAL": "Evaluación Parcial",
    "EVFINAL": "Evaluación Final (ES-FINAL)", "EV DIAG": "Evaluación Diagnóstica (Ficha)"
}

def detectar_tipo_evaluacion(nombre_archivo):
    nombre_upper = nombre_archivo.upper()
    for clave in mapa_equivalencias.keys():
        if clave in nombre_upper:
            return mapa_equivalencias[clave]
    return None

# === FUNCIONES DE PARSING MEJORADAS PARA RANGOS ===
def extraer_mayor_de_rango(texto):
    """
    Extrae el valor mayor de un rango como "9-8" o "9.3-6.7"
    Devuelve el valor mayor como float o None si no encuentra rango
    """
    texto = texto.strip()
    
    # Buscar patrones de rango: 9-8, 9.3-6.7, (9-8), (9.3-6.7), etc.
    patrones_rango = [
        r"\(?(\d+[.,]?\d*)\s*[-–]\s*(\d+[.,]?\d*)\)?",  # 9-8 o 9.3-6.7
        r"\[?(\d+[.,]?\d*)\s*[-–]\s*(\d+[.,]?\d*)\]?",  # [9-8]
        r"(\d+[.,]?\d*)\s*a\s*(\d+[.,]?\d*)",           # 9 a 8
    ]
    
    for patron in patrones_rango:
        match = re.search(patron, texto)
        if match:
            try:
                valor1 = float(match.group(1).replace(",", "."))
                valor2 = float(match.group(2).replace(",", "."))
                return max(valor1, valor2)
            except (ValueError, TypeError):
                continue
    
    # Si no es un rango, intentar extraer un solo número
    match_solo = re.search(r"(\d+[.,]?\d*)", texto)
    if match_solo:
        try:
            return float(match_solo.group(1).replace(",", "."))
        except (ValueError, TypeError):
            pass
    
    return None

def normalizar_puntaje(valor):
    """Normaliza cualquier formato de puntaje a float"""
    if isinstance(valor, (int, float)):
        return float(valor)
    
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
        # Manejar casos como "pts", "puntos", etc.
        if valor.endswith("pts"):
            valor = valor[:-3].strip()
        elif valor.endswith("puntos"):
            valor = valor[:-6].strip()
        
        # Extraer el mayor valor si es un rango
        valor_rango = extraer_mayor_de_rango(valor)
        if valor_rango is not None:
            return valor_rango
        
        # Intentar convertir a float directamente
        try:
            return float(valor)
        except (ValueError, TypeError):
            return None
    
    return None

def extraer_puntaje_y_descripcion_nuevo_formato(texto):
    """
    Versión adaptada para el nuevo formato donde el puntaje puede estar en rangos
    Ejemplo: "Descripción del criterio (9-8)" -> (9.0, "Descripción del criterio")
    """
    texto = texto.strip()
    
    # Primero intentar extraer el puntaje (puede ser un rango)
    puntaje = extraer_mayor_de_rango(texto)
    
    if puntaje is not None:
        # Remover el patrón de rango del texto para obtener la descripción limpia
        descripcion = re.sub(r"\(?\s*\d+[.,]?\d*\s*[-–]\s*\d+[.,]?\d*\s*\)?", "", texto)
        descripcion = re.sub(r"\[?\s*\d+[.,]?\d*\s*[-–]\s*\d+[.,]?\d*\s*\]?", "", descripcion)
        descripcion = re.sub(r"\d+[.,]?\d*\s*a\s*\d+[.,]?\d*", "", descripcion)
        descripcion = descripcion.strip()
        
        # Limpiar paréntesis vacíos y espacios extras
        descripcion = re.sub(r"\(\s*\)", "", descripcion)
        descripcion = re.sub(r"\s+", " ", descripcion).strip()
        
        return puntaje, descripcion
    
    # Si no encuentra patrón de rango, buscar puntaje al final entre paréntesis
    match = re.search(r"(.*?)\s*\((\d+[.,]?\d*)\)\s*$", texto)
    if match:
        descripcion = match.group(1).strip()
        puntaje = normalizar_puntaje(match.group(2))
        return puntaje, descripcion
    
    # Si no encuentra patrón, devolver todo como descripción
    return None, texto

def es_fila_cabecera(texto):
    texto_upper = texto.upper()
    return bool(re.search(r"ASPECTO|CRITERIO|EVALUAR|COMPETENCIAS|LOGRO DESTACADO|LOGRADO|NO LOGRADO", texto_upper))

def es_fila_resumen(texto):
    texto_upper = texto.upper()
    return bool(re.search(r"\d+\s*[-–]\s*\d+|TOTAL|PUNTAJE|NOTA", texto_upper))

# === PROCESAMIENTO DE DOCUMENTO WORD - ADAPTADO PARA RANGOS ===
def procesar_documento(word_file_path):
    try:
        doc = Document(word_file_path)
        evaluacion_data = []
        
        for table in doc.tables:
            # Procesar cada fila de la tabla
            for row in table.rows:
                if len(row.cells) < 4:
                    continue
                    
                celda_0 = row.cells[0].text.strip()
                celda_1 = row.cells[1].text.strip()
                celda_2 = row.cells[2].text.strip()
                celda_3 = row.cells[3].text.strip()
                
                # Saltar filas de cabecera o resumen
                if es_fila_cabecera(celda_0) or es_fila_resumen(celda_0):
                    continue
                    
                if celda_0 and not any(x in celda_0.upper() for x in ["PTS", "PUNTOS", "TOTAL", "NOTA"]):
                    # Extraer puntaje y descripción para cada nivel de logro
                    p_ld, desc_ld = extraer_puntaje_y_descripcion_nuevo_formato(celda_1)
                    p_l, desc_l = extraer_puntaje_y_descripcion_nuevo_formato(celda_2)
                    p_nl, desc_nl = extraer_puntaje_y_descripcion_nuevo_formato(celda_3)
                    
                    # Si no se extrajo puntaje, intentar buscar en el texto
                    if p_ld is None:
                        p_ld = extraer_mayor_de_rango(celda_1)
                    if p_l is None:
                        p_l = extraer_mayor_de_rango(celda_2)
                    if p_nl is None:
                        p_nl = extraer_mayor_de_rango(celda_3)
                    
                    # Si aún no hay puntaje, usar el texto completo como descripción
                    desc_ld = desc_ld if desc_ld else celda_1
                    desc_l = desc_l if desc_l else celda_2
                    desc_nl = desc_nl if desc_nl else celda_3
                    
                    if p_ld is not None and p_l is not None and p_nl is not None:
                        evaluacion_data.append([
                            celda_0,
                            p_ld, "Logro Destacado", desc_ld,
                            p_l, "Logrado", desc_l,
                            p_nl, "No Logrado", desc_nl
                        ])

        if not evaluacion_data:
            error_log.append([word_file_path, "No se encontraron datos válidos en el documento"])
            return None
            
        df = pd.DataFrame(evaluacion_data, columns=[
            "ASPECTO / CRITERIO A EVALUAR", 
            "Puntaje de la calificación", "Título de la calificación", "Descripción de la calificación .1", 
            "Puntaje de la calificación.1", "Título de la calificación .1", "Descripción de la calificación .1.1", 
            "Puntaje de la calificación.2", "Título de la calificación .2", "Descripción de la calificación .2"
        ])
        
        # Validación de suma total (20 puntos) - más flexible para diferentes evaluaciones
        suma_ld = df["Puntaje de la calificación"].sum()
        # Aceptar sumas entre 15-25 para diferentes tipos de evaluación
        if not (15 <= suma_ld <= 25):
            error_log.append([word_file_path, f"La suma de 'Logro Destacado' es inusual: {suma_ld:.2f}"])
            # No retornar None, solo registrar advertencia
            
        return df

    except Exception as e:
        error_log.append([word_file_path, f"Error al procesar el documento: {str(e)}"])
        return None

# === FUNCIONES ADICIONALES ===
def extraer_puntaje_desde_texto(texto):
    texto = texto.strip()
    match = re.search(r"(\d+[.,]?\d*)", texto)
    if match:
        return normalizar_puntaje(match.group(1))
    return None

def procesar_carpetas():
    for root, dirs, files in os.walk(base_folder_path):
        word_files = [f for f in files if f.endswith(".docx") and not f.startswith("~$")]
        if not word_files:
            continue
            
        folder_name = os.path.basename(root)
        output_excel_path = os.path.join(root, f"{folder_name}.xlsx")
        dfs_por_hoja = {}
        
        for filename in word_files:
            word_file_path = os.path.join(root, filename)
            tipo_eval = detectar_tipo_evaluacion(filename)
            
            if not tipo_eval:
                error_log.append([word_file_path, "Tipo de evaluación no reconocido"])
                document_log.append([filename, "Ignorado (tipo no reconocido)"])
                continue
                
            df = procesar_documento(word_file_path)
            if df is not None:
                sheet_name = tipo_eval[:31]
                
                if sheet_name in dfs_por_hoja:
                    counter = 1
                    while f"{sheet_name}_{counter}" in dfs_por_hoja:
                        counter += 1
                    sheet_name = f"{sheet_name}_{counter}"
                
                dfs_por_hoja[sheet_name] = df
                document_log.append([filename, f"Procesado correctamente -> {sheet_name}"])
            else:
                document_log.append([filename, "Procesado con errores"])
        
        if dfs_por_hoja:
            try:
                with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                    for sheet_name, df in dfs_por_hoja.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"✅ Archivo Excel generado: {output_excel_path}")
                
                validation_log.append([
                    output_excel_path,
                    ", ".join(word_files),
                    ", ".join(dfs_por_hoja.keys()),
                    "OK" if len(dfs_por_hoja) == len([f for f in document_log if "correctamente" in f[1]]) else "Algunos archivos no se procesaron"
                ])
            except Exception as e:
                error_log.append([output_excel_path, f"Error al generar Excel: {str(e)}"])

# === EJECUCIÓN Y REPORTES ===
if __name__ == "__main__":
    print("Iniciando procesamiento de documentos...")
    procesar_carpetas()
    
    # Generar reportes
    if error_log:
        pd.DataFrame(error_log, columns=["Archivo", "Error"]).to_excel(
            os.path.join(base_folder_path, "Errores_de_procesamiento.xlsx"), index=False)
        
    if validation_log:
        pd.DataFrame(validation_log, columns=["Archivo Excel", "Archivos Word", "Hojas en Excel", "Observación"]).to_excel(
            os.path.join(base_folder_path, "Resumen_de_validacion.xlsx"), index=False)
        
    if document_log:
        pd.DataFrame(document_log, columns=["Archivo", "Estado"]).to_excel(
            os.path.join(base_folder_path, "Resumen_de_procesamiento.xlsx"), index=False)
    
    print("\nResumen:")
    print(f"- Documentos procesados correctamente: {len([x for x in document_log if 'correctamente' in x[1]])}")
    print(f"- Documentos con errores: {len([x for x in document_log if 'errores' in x[1]])}")
    print(f"- Errores registrados: {len(error_log)}")
    print("\nProceso completado. Revise los archivos de reporte si hay errores.")