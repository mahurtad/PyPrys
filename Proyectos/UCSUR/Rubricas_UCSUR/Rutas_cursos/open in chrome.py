import subprocess
import time

# Ruta de Chrome (ajustar según ubicación en tu PC)
chrome_path = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'

urls = [
    "https://cientificavirtual.cientifica.edu.pe/courses/28036/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/15675/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49735/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/50021/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49580/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/50021/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49735/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49815/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49608/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49311/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49302/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49735/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49255/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49311/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49302/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49875/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/50043/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49608/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/49311/assignments",
    "https://cientificavirtual.cientifica.edu.pe/courses/50043/assignments"
]

# Abre la primera URL en una nueva ventana de Chrome
subprocess.Popen([chrome_path, urls[0]])

# Abre las demás URLs en nuevas pestañas
for url in urls[1:]:
    time.sleep(0.5)
    subprocess.Popen([chrome_path, "--new-tab", url])
