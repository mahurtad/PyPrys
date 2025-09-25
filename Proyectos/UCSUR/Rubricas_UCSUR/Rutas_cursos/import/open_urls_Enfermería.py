import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49745/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49795/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49746/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49500/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49718/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49884/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49779/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50066/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49264/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49373/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50200/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50174/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49589/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49532/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49298/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49849/content_migrations"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
