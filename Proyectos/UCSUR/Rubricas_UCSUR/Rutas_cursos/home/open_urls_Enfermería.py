import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49745",
    "https://cientificavirtual.cientifica.edu.pe/courses/49795",
    "https://cientificavirtual.cientifica.edu.pe/courses/49746",
    "https://cientificavirtual.cientifica.edu.pe/courses/49500",
    "https://cientificavirtual.cientifica.edu.pe/courses/49718",
    "https://cientificavirtual.cientifica.edu.pe/courses/49884",
    "https://cientificavirtual.cientifica.edu.pe/courses/49779",
    "https://cientificavirtual.cientifica.edu.pe/courses/50066",
    "https://cientificavirtual.cientifica.edu.pe/courses/49264",
    "https://cientificavirtual.cientifica.edu.pe/courses/49373",
    "https://cientificavirtual.cientifica.edu.pe/courses/50200",
    "https://cientificavirtual.cientifica.edu.pe/courses/50174",
    "https://cientificavirtual.cientifica.edu.pe/courses/49589",
    "https://cientificavirtual.cientifica.edu.pe/courses/49532",
    "https://cientificavirtual.cientifica.edu.pe/courses/49298",
    "https://cientificavirtual.cientifica.edu.pe/courses/49849"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
