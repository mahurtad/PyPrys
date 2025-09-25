import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49506/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49528/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49735/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50021/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49580/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50021/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49768/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49815/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49608/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49311/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49216/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49715/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50039/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49255/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49302/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49875/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/49770/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50007/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50163/content_migrations",
    "https://cientificavirtual.cientifica.edu.pe/courses/50043/content_migrations"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
