import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49506/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49528/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49735/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50021/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49580/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50021/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49768/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49815/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49608/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49311/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49216/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49715/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50039/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49255/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49302/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49875/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49770/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50007/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50163/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50043/outcomes"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
