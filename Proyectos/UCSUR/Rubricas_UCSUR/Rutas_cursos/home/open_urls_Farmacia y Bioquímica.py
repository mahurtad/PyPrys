import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49604",
    "https://cientificavirtual.cientifica.edu.pe/courses/49168",
    "https://cientificavirtual.cientifica.edu.pe/courses/49620",
    "https://cientificavirtual.cientifica.edu.pe/courses/49696",
    "https://cientificavirtual.cientifica.edu.pe/courses/50125",
    "https://cientificavirtual.cientifica.edu.pe/courses/49878",
    "https://cientificavirtual.cientifica.edu.pe/courses/49595",
    "https://cientificavirtual.cientifica.edu.pe/courses/49799",
    "https://cientificavirtual.cientifica.edu.pe/courses/49720",
    "https://cientificavirtual.cientifica.edu.pe/courses/49566",
    "https://cientificavirtual.cientifica.edu.pe/courses/49287",
    "https://cientificavirtual.cientifica.edu.pe/courses/49988",
    "https://cientificavirtual.cientifica.edu.pe/courses/49776",
    "https://cientificavirtual.cientifica.edu.pe/courses/50118",
    "https://cientificavirtual.cientifica.edu.pe/courses/49714",
    "https://cientificavirtual.cientifica.edu.pe/courses/49633"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
