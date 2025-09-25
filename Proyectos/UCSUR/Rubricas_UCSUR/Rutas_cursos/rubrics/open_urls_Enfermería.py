import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49745/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49795/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49746/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49500/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49718/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49884/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49779/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50066/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49264/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49373/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50200/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50174/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49589/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49532/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49298/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49849/rubrics"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
