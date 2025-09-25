import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49129/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49483/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49525/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49896/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50255/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49712/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50003/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49812/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50213/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49503/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49990/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50170/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49353/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49436/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49621/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50211/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49153/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49508/outcomes"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
