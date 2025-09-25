import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [
        "https://cientificavirtual.cientifica.edu.pe/courses/49604/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49168/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49620/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49696/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50125/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49878/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49595/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49799/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49720/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49566/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49287/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49988/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49776/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/50118/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49714/outcomes",
        "https://cientificavirtual.cientifica.edu.pe/courses/49633/outcomes"
    ]

# Abrir Firefox con la primera URL
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
