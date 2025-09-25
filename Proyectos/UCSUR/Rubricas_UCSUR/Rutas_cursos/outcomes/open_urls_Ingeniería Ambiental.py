import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49946/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49563/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49960/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49668/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49659/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49978/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49284/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49405/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49639/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49665/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49707/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50084/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49830/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50189/outcomes"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
