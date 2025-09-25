import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49946/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49563/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49960/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49668/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49659/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49978/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49284/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49405/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49639/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49665/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49707/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50084/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49830/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50189/content_migrations"
]

# Abrir Firefox con la primera URL en una NUEVA ventana
subprocess.Popen([firefox_path, '-new-window', urls[0]])

# Abrir las demás URLs en pestañas nuevas dentro de la misma ventana
for url in urls[1:]:
    time.sleep(0.5)
    subprocess.Popen([firefox_path, '-new-tab', url])
