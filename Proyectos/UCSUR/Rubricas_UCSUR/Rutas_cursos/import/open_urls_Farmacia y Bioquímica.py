import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49604/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49168/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49620/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49696/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50125/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49878/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49595/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49799/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49720/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49566/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49287/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49988/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49776/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50118/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49714/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49633/content_migrations"
]


# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
