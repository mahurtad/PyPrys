# reporte_evaluaciones_x_curso.py
# -*- coding: utf-8 -*-
import os
import pandas as pd

# ========= CONFIGURACIÓN =========
# Ruta base que contiene la estructura: <ruta_base>/<Carrera>/<Curso>/...
ruta_base = r"C:\Users\manue\OneDrive - Grupo Educad\3. Documentos validados 20252"

# Ruta de salida del reporte (se creará si no existe)
ruta_salida = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\Reporte"
nombre_archivo = "reporte_evaluaciones_por_curso.xlsx"
ruta_reporte = os.path.join(ruta_salida, nombre_archivo)

# Diccionario de mapeo de evaluaciones a siglas estandarizadas
mapeo_evaluaciones = {
    "EC1": "EC1", "EC2": "EC2", "EC3": "EC3", "EC4": "EC4",
    "EF": "EF", "EP": "EP", "ED": "ED",
    "CONTINUA1": "EC1", "CONTINUA2": "EC2", "CONTINUA3": "EC3",
    "CONTINUA4": "EC4", "CONTINUAI": "EC1", "CONTINUAII": "EC2",
    "DIAGNOSTICA": "ED", "FINAL": "EF", "PARCIAL": "EP",
    "EVFINAL": "EF", "EV DIAG": "ED"
}
# =================================


def es_carpeta_curso(ruta_abs: str, base: str) -> bool:
    """
    Devuelve True si 'ruta_abs' corresponde exactamente al nivel <base>/<Carrera>/<Curso>.
    """
    rel = os.path.relpath(ruta_abs, base)
    if rel == ".":
        return False
    partes = rel.split(os.sep)
    return len(partes) == 2  # carrera/curso


def carrera_y_curso_desde_root(ruta_abs: str, base: str):
    """
    Obtiene (carrera, curso) tomando las dos primeras partes relativas a 'base'.
    Si no hay suficientes niveles, devuelve (None, None).
    """
    rel = os.path.relpath(ruta_abs, base)
    if rel == ".":
        return (None, None)
    partes = rel.split(os.sep)
    if len(partes) < 2:
        return (None, None)
    return (partes[0], partes[1])


def normaliza_nombre(s: str) -> str:
    return (s or "").strip()


def main():
    if not os.path.isdir(ruta_base):
        raise FileNotFoundError(f"No existe la ruta_base: {ruta_base}")

    # Estructuras de acumulación por (carrera, curso)
    evaluaciones_por_curso = {}  # (carrera, curso) -> set(evaluaciones)
    archivos_por_curso = {}      # (carrera, curso) -> conteo total de archivos
    ruta_curso_map = {}          # (carrera, curso) -> ruta absoluta del curso

    # Primer pase: recorrer toda la estructura y asignar cada root a su (carrera, curso)
    for root, dirs, files in os.walk(ruta_base):
        carrera, curso = carrera_y_curso_desde_root(root, ruta_base)
        if carrera is None or curso is None:
            continue

        key = (normaliza_nombre(carrera), normaliza_nombre(curso))

        # Inicializar estructuras
        if key not in evaluaciones_por_curso:
            evaluaciones_por_curso[key] = set()
        if key not in archivos_por_curso:
            archivos_por_curso[key] = 0
        if key not in ruta_curso_map:
            # Guardamos la ruta del directorio de curso (nivel exacto carrera/curso)
            ruta_curso_map[key] = os.path.join(ruta_base, carrera, curso)

        # Contar archivos en este nivel
        archivos_por_curso[key] += len(files)

        # Detectar evaluaciones por nombres de subcarpetas y archivos
        for nombre in list(dirs) + list(files):
            nombre_mayus = nombre.upper()
            for clave, sigla in mapeo_evaluaciones.items():
                if clave in nombre_mayus:
                    evaluaciones_por_curso[key].add(sigla)

    # Construir filas del reporte para todos los cursos, incluso si no hay archivos/evaluaciones
    filas = []
    for key in sorted(evaluaciones_por_curso.keys()):
        carrera, curso = key
        evals = evaluaciones_por_curso[key]
        cant_arch = archivos_por_curso.get(key, 0)
        ruta_curso = ruta_curso_map.get(key, os.path.join(ruta_base, carrera, curso))
        filas.append({
            "Carrera": carrera,
            "Curso": curso,
            "Cantidad de Evaluaciones": len(evals),
            "Evaluaciones": ", ".join(sorted(evals)) if evals else "",
            "Total de archivos (incluye subcarpetas)": cant_arch,
            "Carpeta sin archivos": "Sí" if cant_arch == 0 else "No",
            "Ruta de la carpeta del Curso": ruta_curso
        })

    # También contemplar el caso en que existan carreras/curso sin haber sido
    # visitados (muy raro) pero que sí están físicamente — esto se cubre de forma
    # natural con os.walk, así que normalmente no hace falta nada extra.

    df = pd.DataFrame(filas).sort_values(["Carrera", "Curso"], kind="stable")

    # Crear carpeta de salida si no existe
    os.makedirs(ruta_salida, exist_ok=True)

    # Guardar reporte
    df.to_excel(ruta_reporte, index=False)
    print(f"Reporte generado correctamente en:\n{ruta_reporte}")
    print(f"Total de cursos listados: {len(df)}")


if __name__ == "__main__":
    main()
