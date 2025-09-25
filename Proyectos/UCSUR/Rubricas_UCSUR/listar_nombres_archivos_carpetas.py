import os
import pandas as pd

def obtener_nombres_carpetas(ruta_base):
    carreras = []
    cursos = []
    
    if not os.path.exists(ruta_base):
        print(f"La ruta {ruta_base} no existe.")
        return None
    
    for carrera in os.listdir(ruta_base):
        carrera_path = os.path.join(ruta_base, carrera)
        if os.path.isdir(carrera_path):
            for curso in os.listdir(carrera_path):
                curso_path = os.path.join(carrera_path, curso)
                if os.path.isdir(curso_path):
                    carreras.append(carrera)
                    cursos.append(curso)
    
    return pd.DataFrame({"Carrera": carreras, "Curso": cursos})

def guardar_en_excel(df, nombre_archivo):
    df.to_excel(nombre_archivo, index=False)
    print(f"Archivo guardado como {nombre_archivo}")

# Ruta base (ajustar según necesidad)
ruta_base = r"G:\My Drive\Rubricas UCSUR\3. Documentos validados"

# Generar DataFrame y guardar en Excel
df_carpetas = obtener_nombres_carpetas(ruta_base)
if df_carpetas is not None:
    guardar_en_excel(df_carpetas, "lista_carpetas.xlsx")
