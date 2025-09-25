import os

# Ruta base donde están los cursos
ruta_destino_base = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\cursos"

# Nombres de carpetas de curso
nombres_cursos = [
    "Biologia",
    "Bioquimica",
    "Estadistica General",
    "Fisica",
    "Fisica I",
    "Quimica general",
    "Química Orgánica"
]

# Nombre actual y nuevo del ZIP
nombre_antiguo = "Lecturas o bibliografía obligatoria.zip"
nombre_nuevo = "Lecturas obligatorias.zip"

# Recorrer cada curso
for nombre_curso in nombres_cursos:
    carpeta_curso = os.path.join(ruta_destino_base, nombre_curso)
    zip_antiguo = os.path.join(carpeta_curso, nombre_antiguo)
    zip_nuevo = os.path.join(carpeta_curso, nombre_nuevo)

    if os.path.exists(zip_antiguo):
        os.rename(zip_antiguo, zip_nuevo)
        print(f"Renombrado: {zip_antiguo} -> {zip_nuevo}")
    else:
        print(f"No encontrado: {zip_antiguo}")

print("Proceso completado.")
