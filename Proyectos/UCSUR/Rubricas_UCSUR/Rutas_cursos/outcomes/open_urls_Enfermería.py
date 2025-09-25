import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49745/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49795/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49746/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49500/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49718/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49884/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49779/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50066/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49264/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49373/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50200/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50174/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49589/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49532/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49298/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49849/outcomes"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
