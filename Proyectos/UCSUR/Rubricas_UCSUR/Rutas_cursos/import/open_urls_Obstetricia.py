import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49129/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49483/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49525/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49896/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50255/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49712/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50003/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49812/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50213/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49503/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49990/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50170/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49353/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49436/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49621/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50211/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49153/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49508/content_migrations"
]

# Abrir Firefox con la primera URL en una NUEVA ventana
subprocess.Popen([firefox_path, '-new-window', urls[0]])

# Abrir las demás URLs en pestañas nuevas dentro de la misma ventana
for url in urls[1:]:
    time.sleep(0.5)
    subprocess.Popen([firefox_path, '-new-tab', url])