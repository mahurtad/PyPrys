import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/49129",
    "https://cientificavirtual.cientifica.edu.pe/courses/49483",
    "https://cientificavirtual.cientifica.edu.pe/courses/49525",
    "https://cientificavirtual.cientifica.edu.pe/courses/49896",
    "https://cientificavirtual.cientifica.edu.pe/courses/50255",
    "https://cientificavirtual.cientifica.edu.pe/courses/49712",
    "https://cientificavirtual.cientifica.edu.pe/courses/50003",
    "https://cientificavirtual.cientifica.edu.pe/courses/49812",
    "https://cientificavirtual.cientifica.edu.pe/courses/50213",
    "https://cientificavirtual.cientifica.edu.pe/courses/49503",
    "https://cientificavirtual.cientifica.edu.pe/courses/49990",
    "https://cientificavirtual.cientifica.edu.pe/courses/50170",
    "https://cientificavirtual.cientifica.edu.pe/courses/49353",
    "https://cientificavirtual.cientifica.edu.pe/courses/49436",
    "https://cientificavirtual.cientifica.edu.pe/courses/49621",
    "https://cientificavirtual.cientifica.edu.pe/courses/50211",
    "https://cientificavirtual.cientifica.edu.pe/courses/49153",
    "https://cientificavirtual.cientifica.edu.pe/courses/49508"
]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
