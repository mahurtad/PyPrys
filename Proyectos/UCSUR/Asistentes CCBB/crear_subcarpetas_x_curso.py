import os

# Lista de cursos
cursos = [
    "Biologia",
    "Quimica general",
    "Química Orgánica",
    "Matemática",
    "Matematica general",
    "Matematica I",
    "Matematica II",
    "Álgebra",
    "Matematica III",
    "Bioquimica",
    "Estadistica General",
    "Fisica",
    "Fisica I"
]

# Lista de subcarpetas a crear dentro de cada curso
subcarpetas = [
    "Sílabo",
    "Ficha de actividades evaluadas",
    "Presentaciones",
    "Videos del curso",
    "Lecturas o bibliografía obligatoria",
    "Materiales adicionales"
]

# Ruta base
ruta_base = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"

# Crear carpetas y subcarpetas
for curso in cursos:
    ruta_curso = os.path.join(ruta_base, curso)
    try:
        os.makedirs(ruta_curso, exist_ok=True)
        for sub in subcarpetas:
            ruta_sub = os.path.join(ruta_curso, sub)
            os.makedirs(ruta_sub, exist_ok=True)
        print(f"Estructura creada para: {curso}")
    except Exception as e:
        print(f"Error al crear estructura para {curso}: {e}")
