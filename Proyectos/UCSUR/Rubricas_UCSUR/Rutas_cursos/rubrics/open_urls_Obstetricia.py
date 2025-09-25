import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49129/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49483/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49525/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49896/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50255/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49712/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50003/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49812/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50213/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49503/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49990/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50170/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49353/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49436/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49621/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50211/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49153/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49508/rubrics"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
