import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49506",
        "https://cientificavirtual.cientifica.edu.pe/courses/49528",
        "https://cientificavirtual.cientifica.edu.pe/courses/49735",
        "https://cientificavirtual.cientifica.edu.pe/courses/50021",
        "https://cientificavirtual.cientifica.edu.pe/courses/49580",
        "https://cientificavirtual.cientifica.edu.pe/courses/50021",
        "https://cientificavirtual.cientifica.edu.pe/courses/49768",
        "https://cientificavirtual.cientifica.edu.pe/courses/49815",
        "https://cientificavirtual.cientifica.edu.pe/courses/49608",
        "https://cientificavirtual.cientifica.edu.pe/courses/49311",
        "https://cientificavirtual.cientifica.edu.pe/courses/49216",
        "https://cientificavirtual.cientifica.edu.pe/courses/49715",    
        "https://cientificavirtual.cientifica.edu.pe/courses/50039",
        "https://cientificavirtual.cientifica.edu.pe/courses/49255",
        "https://cientificavirtual.cientifica.edu.pe/courses/49302",
        "https://cientificavirtual.cientifica.edu.pe/courses/49875",
        "https://cientificavirtual.cientifica.edu.pe/courses/49770",
        "https://cientificavirtual.cientifica.edu.pe/courses/50007",
        "https://cientificavirtual.cientifica.edu.pe/courses/50163",
        "https://cientificavirtual.cientifica.edu.pe/courses/50043"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
