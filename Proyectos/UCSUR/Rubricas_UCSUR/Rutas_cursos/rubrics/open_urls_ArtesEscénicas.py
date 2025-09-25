import subprocess
import time

# Ruta de Firefox (ajustar si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista de URLs
urls = [  
    "https://cientificavirtual.cientifica.edu.pe/courses/49506/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49528/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49735/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49578/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49580/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/50021/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49768/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49815/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49608/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49311/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49216/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49715/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/50039/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49255/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49302/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/49875/rubrics", 
    "https://cientificavirtual.cientifica.edu.pe/courses/49770/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/50007/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/50163/rubrics",
    "https://cientificavirtual.cientifica.edu.pe/courses/50043/rubrics"
]
# Abrir Firefox con la primera UR
subprocess.Popen([firefox_path, urls[0]])

# Abrir las demás URLs en pestañas nuevas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa breve
    subprocess.Popen([firefox_path, "-new-tab", url])
