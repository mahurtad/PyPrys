import subprocess
import time

# Ruta de Firefox (ajusta si es necesario)
firefox_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"

# Lista base de cursos
curso_ids = [
    "58736", "59385", "59681", "59043", "59128", "59690", "59683", "59701",
    "59695", "59685", "59692", "59295", "59666", "59109", "59198", "59245",
    "59699", "59658", "58884", "58890", "59294", "59275"
]

# Generar las URLs completas
urls = [f"https://cientificavirtual.cientifica.edu.pe/courses/{id}/content_migrations" for id in curso_ids]

# Abrir la primera en nueva ventana
subprocess.Popen([firefox_path, '-new-window', urls[0]])

# Abrir el resto en nuevas pestañas
for url in urls[1:]:
    time.sleep(0.5)  # Pausa para evitar saturar
    subprocess.Popen([firefox_path, '-new-tab', url])
