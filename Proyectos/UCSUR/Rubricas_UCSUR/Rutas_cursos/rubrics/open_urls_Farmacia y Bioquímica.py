import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49604/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49168/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49620/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49696/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50125/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49878/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49595/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49799/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49720/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49566/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49287/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49988/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49776/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/50118/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49714/rubrics",
        "https://cientificavirtual.cientifica.edu.pe/courses/49633/rubrics"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
