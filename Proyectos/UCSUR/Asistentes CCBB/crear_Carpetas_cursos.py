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

# Ruta base donde se crearán las carpetas
ruta_base = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"

# Crear las carpetas
for curso in cursos:
    ruta_carpeta = os.path.join(ruta_base, curso)
    try:
        os.makedirs(ruta_carpeta, exist_ok=True)
        print(f"Carpeta creada o ya existente: {ruta_carpeta}")
    except Exception as e:
        print(f"Error al crear carpeta {curso}: {e}")
