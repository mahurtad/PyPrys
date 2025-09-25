import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49946",
    "https://cientificavirtual.cientifica.edu.pe/courses/49563",
    "https://cientificavirtual.cientifica.edu.pe/courses/49960",
    "https://cientificavirtual.cientifica.edu.pe/courses/49668",
    "https://cientificavirtual.cientifica.edu.pe/courses/49659",
    "https://cientificavirtual.cientifica.edu.pe/courses/49978",
    "https://cientificavirtual.cientifica.edu.pe/courses/49284",
    "https://cientificavirtual.cientifica.edu.pe/courses/49405",
    "https://cientificavirtual.cientifica.edu.pe/courses/49639",
    "https://cientificavirtual.cientifica.edu.pe/courses/49665",
    "https://cientificavirtual.cientifica.edu.pe/courses/49707",
    "https://cientificavirtual.cientifica.edu.pe/courses/50084",
    "https://cientificavirtual.cientifica.edu.pe/courses/49830",
    "https://cientificavirtual.cientifica.edu.pe/courses/50189"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
